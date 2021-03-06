/* global JsonRoutes */
if (Meteor.isServer) {
  // todo XXX once we have a real API in place, move that route there
  // todo XXX also  share the route definition between the client and the server
  // so that we could use something like
  // `ApiRoutes.path('boards/exportExcel', boardId)``
  // on the client instead of copy/pasting the route path manually between the
  // client and the server.
  /**
   * @operation exportExcel
   * @tag Boards
   *
   * @summary This route is used to export the board Excel.
   *
   * @description If user is already logged-in, pass loginToken as param
   * "authToken": '/api/boards/:boardId/exportExcel?authToken=:token'
   *
   * See https://blog.kayla.com.au/server-side-route-authentication-in-meteor/
   * for detailed explanations
   *
   * @param {string} boardId the ID of the board we are exporting
   * @param {string} authToken the loginToken
   */
  var Excel = require('exceljs');
  Picker.route('/vag/api/boards/:boardId/exportExcel',
    function(params, req, res, next) {
      const boardId = params.boardId;
      let user = null;
      console.log('Excel');

      const loginToken = params.query.authToken;
      if (loginToken) {
        const hashToken = Accounts._hashLoginToken(loginToken);
        user = Meteor.users.findOne({
          'services.resume.loginTokens.hashedToken': hashToken,
        });
      } else if (!Meteor.settings.public.sandstorm) {
        Authentication.checkUserId(req.userId);
        user = Users.findOne({
          _id: req.userId,
          isAdmin: true
        });
      }
      const exporterExcel = new ExporterExcel(boardId);
      if (exporterExcel.canExport(user)) {
        exporterExcel.build(res);
        //);
      } else {
        res.end('导!出!异!常!');
      }
    });
}

// exporter maybe is broken since Gridfs introduced, add fs and path

export class ExporterExcel {
  constructor(boardId) {
    this._boardId = boardId;
  }

  build(res) {
    const fs = Npm.require('fs');
    const os = Npm.require('os');
    const path = Npm.require('path');

    const byBoard = {
      boardId: this._boardId
    };
    const byBoardNoLinked = {
      boardId: this._boardId,
      linkedId: {
        $in: ['', null]
      },
    };
    // we do not want to retrieve boardId in related elements
    const noBoardId = {
      fields: {
        boardId: 0,
      },
    };
    const result = {
      _format: 'wekan-board-1.0.0',
    };
    _.extend(
      result,
      Boards.findOne(this._boardId, {
        fields: {
          stars: 0,
        },
      }),
    );
    result.lists = Lists.find(byBoard, noBoardId).fetch();
    result.cards = Cards.find(byBoardNoLinked, noBoardId).fetch();
    result.swimlanes = Swimlanes.find(byBoard, noBoardId).fetch();
    result.customFields = CustomFields.find({
      boardIds: {
        $in: [this.boardId]
      }
    }, {
      fields: {
        boardId: 0
      }
    }, ).fetch();
    result.comments = CardComments.find(byBoard, noBoardId).fetch();
    result.activities = Activities.find(byBoard, noBoardId).fetch();
    result.rules = Rules.find(byBoard, noBoardId).fetch();
    result.checklists = [];
    result.checklistItems = [];
    result.subtaskItems = [];
    result.triggers = [];
    result.actions = [];
    result.cards.forEach(card => {
      result.checklists.push(
        ...Checklists.find({
          cardId: card._id,
        }).fetch(),
      );
      result.checklistItems.push(
        ...ChecklistItems.find({
          cardId: card._id,
        }).fetch(),
      );
      result.subtaskItems.push(
        ...Cards.find({
          parentId: card._id,
        }).fetch(),
      );
    });
    result.rules.forEach(rule => {
      result.triggers.push(
        ...Triggers.find({
            _id: rule.triggerId,
          },
          noBoardId,
        ).fetch(),
      );
      result.actions.push(
        ...Actions.find({
            _id: rule.actionId,
          },
          noBoardId,
        ).fetch(),
      );
    });

    // [Old] for attachments we only export IDs and absolute url to original doc
    // [New] Encode attachment to base64
    const getBase64Data = function(doc, callback) {
      let buffer = new Buffer(0);
      // callback has the form function (err, res) {}
      const tmpFile = path.join(
        os.tmpdir(),
        `tmpexport${process.pid}${Math.random()}`,
      );
      const tmpWriteable = fs.createWriteStream(tmpFile);
      const readStream = doc.createReadStream();
      readStream.on('data', function(chunk) {
        buffer = Buffer.concat([buffer, chunk]);
      });
      readStream.on('error', function(err) {
        callback(err, null);
      });
      readStream.on('end', function() {
        // done
        fs.unlink(tmpFile, () => {
          //ignored
        });
        callback(null, buffer.toString('base64'));
      });
      readStream.pipe(tmpWriteable);
    };
    //const getBase64DataSync = Meteor.wrapAsync(getBase64Data);
    //result.attachments = Attachments.find(byBoard)
     // .fetch()
      //.map(attachment => {
       // return {
        //  _id: attachment._id,
         // cardId: attachment.cardId,
          // url: FlowRouter.url(attachment.url()),
          //file: getBase64DataSync(attachment),
        //  name: attachment.original.name,
         // type: attachment.original.type,
      //  };
      //});

    // we also have to export some user data - as the other elements only
    // include id but we have to be careful:
    // 1- only exports users that are linked somehow to that board
    // 2- do not export any sensitive information
    const users = {};
    result.members.forEach(member => {
      users[member.userId] = true;
    });
    result.lists.forEach(list => {
      users[list.userId] = true;
    });
    result.cards.forEach(card => {
      users[card.userId] = true;
      if (card.members) {
        card.members.forEach(memberId => {
          users[memberId] = true;
        });
      }
    });
    result.comments.forEach(comment => {
      users[comment.userId] = true;
    });
    result.activities.forEach(activity => {
      users[activity.userId] = true;
    });
    result.checklists.forEach(checklist => {
      users[checklist.userId] = true;
    });
    const byUserIds = {
      _id: {
        $in: Object.getOwnPropertyNames(users),
      },
    };
    // we use whitelist to be sure we do not expose inadvertently
    // some secret fields that gets added to User later.
    const userFields = {
      fields: {
        _id: 1,
        username: 1,
        'profile.fullname': 1,
        'profile.initials': 1,
        'profile.avatarUrl': 1,
      },
    };
    result.users = Users.find(byUserIds, userFields)
      .fetch()
      .map(user => {
        // user avatar is stored as a relative url, we export absolute
        if ((user.profile || {}).avatarUrl) {
          user.profile.avatarUrl = FlowRouter.url(user.profile.avatarUrl);
        }
        return user;
      });

    var jdata = result;
    //init exceljs workbook
    var workbook = new Excel.Workbook();
    workbook.creator = 'wekan';
    workbook.lastModifiedBy = 'wekan';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.lastPrinted = new Date();
    var filename = jdata.title + "-看板导出.xlsx";
    //init worksheet
    var worksheet = workbook.addWorksheet(jdata.title, {
      properties: {
        tabColor: {
          argb: 'FFC0000'
        }
      },
      pageSetup: {
        paperSize: 9,
        orientation: 'landscape'
      }
    });
    //get worksheet
    var ws = workbook.getWorksheet(jdata.title);
    ws.properties.defaultRowHeight = 20;
    //init columns
    ws.columns = [{
        key: 'a',
        width: 7
      }, {
        key: 'b',
        width: 16
      }, {
        key: 'c',
        width: 7
      }, {
        key: 'd',
        width: 14,
        style: {
          font: {
            name: '宋体',
            size: '10'
          },
          numFmt: 'yyyy/mm/dd hh:mm:ss'
        }
      }, {
        key: 'e',
        width: 14,
        style: {
          font: {
            name: '宋体',
            size: '10'
          },
          numFmt: 'yyyy/mm/dd hh:mm:ss'
        }
      }, {
        key: 'f',
        width: 10
      },
      {
        key: 'g',
        width: 10
      },
      {
        key: 'h',
        width: 18
      }
    ];

    //add title line
    ws.mergeCells('A1:H1');
    ws.getCell('A1').value = jdata.title;
    ws.getCell('A1').style = {
      font: {
        name: '宋体',
        size: '20'
      }
    };
    ws.getCell('A1').alignment = {
      vertical: 'middle',
      horizontal: 'center'
    };
    ws.getRow(1).height = 40;
    //get member info
    var jmem = "";
    var jmeml = {};
    for (var i in jdata.users) {
      jmem = jmem + jdata.users[i].profile.fullname + ",";
      jmeml[jdata.users[i]._id] = jdata.users[i].profile.fullname;
    }
    jmem = jmem.substr(0, jmem.length - 1);
    //get kanban list info
    var jlist = {};
    for (var k in jdata.lists) {
      jlist[jdata.lists[k]._id] = jdata.lists[k].title;
    }
    //get kanban label info
    var jlabel= {};
    for (var k in jdata.labels) {
      jlabel[jdata.labels[k]._id] = jdata.labels[k].name;
    }
    //add data +8 hours
    function add8hours(jdate) {
      var curdate = new Date(jdate);
      return new Date(curdate.setHours(curdate.getHours() + 8));
    }
    //add blank row
    ws.addRow().values = ['', '', '', '', '', '', '', ''];
    //add kanban info
    ws.addRow().values = ['创建时间', add8hours(jdata.createdAt), '更新时间', add8hours(jdata.modifiedAt), '成员', jmem];
    ws.getRow(3).font = {
      name: '宋体',
      size: 10,
      bold: true
    };
    ws.getCell('B3').style = {
      font: {
        name: '宋体',
        size: '10',
        bold: true
      },
      numFmt: 'yyyy/mm/dd hh:mm:ss'
    };
    //cell center
    function cellCenter(cellno) {
      ws.getCell(cellno).alignment = {
        vertical: 'middle',
        horizontal: 'center',
        wrapText: true
      };
    }
    cellCenter('A3');
    cellCenter('B3');
    cellCenter('C3');
    cellCenter('D3');
    cellCenter('E3');
    cellCenter('F3');
    ws.getRow(3).height = 20;
    //all border
    function allBorder(cellno) {
      ws.getCell(cellno).border = {
        top: {
          style: 'thin'
        },
        left: {
          style: 'thin'
        },
        bottom: {
          style: 'thin'
        },
        right: {
          style: 'thin'
        }
      };
    }
    allBorder('A3');
    allBorder('B3');
    allBorder('C3');
    allBorder('D3');
    allBorder('E3');
    allBorder('F3');
    //add blank row
    ws.addRow().values = ['', '', '', '', '', '', '', '', ''];
    //add card title
    ws.addRow().values = ['编号', '标题', '创建人', '创建时间', '更新时间', '列表', '成员', '描述', '标签'];
    ws.getRow(5).height = 20;
    allBorder('A5');
    allBorder('B5');
    allBorder('C5');
    allBorder('D5');
    allBorder('E5');
    allBorder('F5');
    allBorder('G5');
    allBorder('H5');
    allBorder('I5');
    cellCenter('A5');
    cellCenter('B5');
    cellCenter('C5');
    cellCenter('D5');
    cellCenter('E5');
    cellCenter('F5');
    cellCenter('G5');
    cellCenter('H5');
    cellCenter('I5');
    ws.getRow(5).font = {
      name: '宋体',
      size: 12,
      bold: true
    };
    //add blank row
    //add card info
    for (var i in jdata.cards) {
      var jcard = jdata.cards[i]
      //get member info
      var jcmem = "";
      for (var j in jcard.members) {
        jcmem += jmeml[jcard.members[j]];
        jcmem += " ";
      }
      //get card label info
      var jclabel ="";
      for (var j in jcard.labelIds) {
        jclabel += jlabel[jcard.labelIds[j]];
        jclabel += " ";
      }
//      console.log(jclabel);

      //add card detail
      var t = Number(i) + 1;
      ws.addRow().values = [t.toString(), jcard.title, jmeml[jcard.userId], add8hours(jcard.createdAt), add8hours(jcard.dateLastActivity), jlist[jcard.listId], jcmem, jcard.description, jclabel];
      var y = Number(i) + 6;
      //ws.getRow(y).height = 25;
      allBorder('A' + y);
      allBorder('B' + y);
      allBorder('C' + y);
      allBorder('D' + y);
      allBorder('E' + y);
      allBorder('F' + y);
      allBorder('G' + y);
      allBorder('H' + y);
      allBorder('I' + y);
      cellCenter('A' + y);
      ws.getCell('B' + y).alignment = {
        wrapText: true
      };
      ws.getCell('H' + y).alignment = {
        wrapText: true
      };
      ws.getCell('I' + y).alignment = {
        wrapText: true
      };
    }
    //    var exporte=new Stream;
    workbook.xlsx.write(res)
      .then(function() {});
    //     return exporte;
  }


  canExport(user) {
    const board = Boards.findOne(this._boardId);
    return board && board.isVisibleBy(user);
  }


}
