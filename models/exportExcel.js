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
  JsonRoutes.add('get', '/api/boards/:boardId/exportExcel', function(req, res) {
    const boardId = req.params.boardId;
    let user = null;

    const loginToken = req.query.authToken;
    if (loginToken) {
      const hashToken = Accounts._hashLoginToken(loginToken);
      user = Meteor.users.findOne({
        'services.resume.loginTokens.hashedToken': hashToken,
      });
    } else if (!Meteor.settings.public.sandstorm) {
      Authentication.checkUserId(req.userId);
      user = Users.findOne({ _id: req.userId, isAdmin: true });
    }

    const exporter = new Exporter(boardId);
    if (exporter.canExport(user)) {
      JsonRoutes.sendResult(res, {
        code: 200,
        data: exporter.build(),
      });
    } else {
      // we could send an explicit error message, but on the other hand the only
      // way to get there is by hacking the UI so let's keep it raw.
      JsonRoutes.sendResult(res, 403);
    }
  });
}

// exporter maybe is broken since Gridfs introduced, add fs and path

export class Exporter {
  constructor(boardId) {
    this._boardId = boardId;
  }

  build() {
    const fs = Npm.require('fs');
    const os = Npm.require('os');
    const path = Npm.require('path');

    const byBoard = { boardId: this._boardId };
    const byBoardNoLinked = {
      boardId: this._boardId,
      linkedId: { $in: ['', null] },
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
    result.customFields = CustomFields.find(
      { boardIds: { $in: [this.boardId] } },
      { fields: { boardId: 0 } },
    ).fetch();
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
        ...Triggers.find(
          {
            _id: rule.triggerId,
          },
          noBoardId,
        ).fetch(),
      );
      result.actions.push(
        ...Actions.find(
          {
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
    const getBase64DataSync = Meteor.wrapAsync(getBase64Data);
    result.attachments = Attachments.find(byBoard)
      .fetch()
      .map(attachment => {
        return {
          _id: attachment._id,
          cardId: attachment.cardId,
          // url: FlowRouter.url(attachment.url()),
          file: getBase64DataSync(attachment),
          name: attachment.original.name,
          type: attachment.original.type,
        };
      });

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
      }
    });
    //get worksheet
    var ws = workbook.getWorksheet(jdata.title);
    //init columns
    ws.columns = [{
      key: 'a'
    }, {
      key: 'b'
    }, {
      key: 'c'
    }, {
      key: 'd'
    }, {
      key: 'e'
    }, {
      key: 'f'
    }]
    //add title line
    ws.mergeCells('A1:H1');
    ws.getCell('A1').value = jdata.title;
    ws.getCell('A1').alignment = {
      vertical: 'middle',
      horizontal: 'center'
    };
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
    //add data +8 hours
    function add8hours(jdate) {
      curdate = new Date(jdate);
      return new Date(curdate.setHours(curdate.getHours() + 8));
    }
    //add kanban info
    ws.addRow().values = ['创建时间', add8hours(jdata.createdAt), '最近更新时间', add8hours(jdata.modifiedAt), '成员', jmem]
    //add card title
    ws.addRow().values = ['编号', '标题', '描述', '创建人', '创建时间', '更新时间', '列表', '成员', '']
    //add card info
    for (var i in jdata.cards) {
      jcard = jdata.cards[i]
      //get member info
      var jcmem = "";
      for (var j in jcard.members) {
        jcmem = jcmem + jmeml[jcard.members[j]];
      }
      //add card detail
      t = Number(i) + 1
      ws.addRow().values = [t.toString(), jcard.title, jcard.discription, jmeml[jcard.userId], add8hours(jcard.createdAt), add8hours(jcard.dateLastActivity), jlist[jcard.listId], jcmem];
    }
    var exporte = Buffer();
    workbook.xlsx.writeBuffer()
      .then(function(exporte) {
        // done
      });
    return exporte;

//    return result;
  }


  canExport(user) {
    const board = Boards.findOne(this._boardId);
    return board && board.isVisibleBy(user);
  }


}
