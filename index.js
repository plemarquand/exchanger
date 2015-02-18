var path = require('path');
var moment = require('moment');
var crypto = require('crypto');
var xml2js = require('xml2js');
var Bacon = require('baconjs');

exports.client = null;

exports.initialize = function(settings, callback) {
  var soap = require('soap');
  // TODO: Handle different locations of where the asmx lives.
  var endpoint = 'https://' + path.join(settings.url, 'EWS/Exchange.asmx');
  var url = path.join(__dirname, 'Services.wsdl');

  soap.createClient(url, {}, function(err, client) {
    if (err) {
      return callback(err);
    }
    if (!client) {
      return callback(new Error('Could not create client'));
    }

    exports.client = client;
    exports.client.setSecurity(new soap.NtlmSecurity(settings.username, settings.password));

    return callback(null);
  }, endpoint);
};

exports.getEmails = function(folderName, limit, callback) {
  if (typeof(folderName) === "function") {
    callback = folderName;
    folderName = 'inbox';
    limit = 10;
  }
  if (typeof(limit) === "function") {
    callback = limit;
    limit = 10;
  }
  if (!exports.client) {
    return callback(new Error('Call initialize()'));
  }

  var soapRequest =
    '<tns:FindItem Traversal="Shallow" xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
      '<tns:ItemShape>' +
        '<t:BaseShape>IdOnly</t:BaseShape>' +
        '<t:AdditionalProperties>' +
          '<t:FieldURI FieldURI="item:ItemId"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="item:ConversationId"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="message:ReplyTo"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="message:ToRecipients"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="message:CcRecipients"></t:FieldURI>' +
          // '<t:FieldURI FieldURI="message:BccRecipients"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:DateTimeCreated"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:DateTimeSent"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:HasAttachments"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:Size"></t:FieldURI>' +
          '<t:FieldURI FieldURI="message:From"></t:FieldURI>' +
          '<t:FieldURI FieldURI="message:IsRead"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:Importance"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:Subject"></t:FieldURI>' +
          '<t:FieldURI FieldURI="item:DateTimeReceived"></t:FieldURI>' +
        '</t:AdditionalProperties>' +
      '</tns:ItemShape>' +
      '<tns:IndexedPageItemView BasePoint="Beginning" Offset="0" MaxEntriesReturned="10"></tns:IndexedPageItemView>' +
      '<tns:ParentFolderIds>' +
        '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' +
      '</tns:ParentFolderIds>' +
    '</tns:FindItem>';

  exports.client.FindItem(soapRequest, function(err, result, body) {
    if (err) {
      return callback(err);
    }

    var parser = new xml2js.Parser();

    parser.parseString(body, function(err, result) {
      var responseCode = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:ResponseCode'];

      if (responseCode !== 'NoError') {
        return callback(new Error(responseCode));
      }

      var rootFolder = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:RootFolder'];

      var emails = [];
      rootFolder['t:Items']['t:Message'].forEach(function(item, idx) {
        var md5hasher = crypto.createHash('md5');
        md5hasher.update(item['t:Subject'] + item['t:DateTimeSent']);
        var hash = md5hasher.digest('hex');

        var itemId = {
          id: item['t:ItemId']['@'].Id,
          changeKey: item['t:ItemId']['@'].ChangeKey
        };

        var dateTimeReceived = item['t:DateTimeReceived'];

        emails.push({
          id: itemId.id + '|' + itemId.changeKey,
          hash: hash,
          subject: item['t:Subject'],
          dateTimeReceived: moment(dateTimeReceived).format("MM/DD/YYYY, h:mm:ss A"),
          size: item['t:Size'],
          importance: item['t:Importance'],
          hasAttachments: (item['t:HasAttachments'] === 'true'),
          from: item['t:From']['t:Mailbox']['t:Name'],
          isRead: (item['t:IsRead'] === 'true'),
          meta: {
            itemId: itemId
          }
        });
      });

      callback(null, emails);
    });
  });
};


exports.getEmail = function(itemId, callback) {
  if (!exports.client) {
    return callback(new Error('Call initialize()'));
  }
  if ((!itemId.id || !itemId.changeKey) && itemId.indexOf('|') > 0) {
    var s = itemId.split('|');

    itemId = {
      id: itemId.split('|')[0],
      changeKey: itemId.split('|')[1]
    };
  }

  if (!itemId.id || !itemId.changeKey) {
    return callback(new Error('Id is not correct.'));
  }

  var soapRequest =
    '<tns:GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
      '<tns:ItemShape>' +
        '<t:BaseShape>Default</t:BaseShape>' +
        '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
      '</tns:ItemShape>' +
      '<tns:ItemIds>' +
        '<t:ItemId Id="' + itemId.id + '" ChangeKey="' + itemId.changeKey + '" />' +
      '</tns:ItemIds>' +
    '</tns:GetItem>';

  exports.client.GetItem(soapRequest, function(err, result, body) {
    if (err) {
      return callback(err);
    }

    var parser = new xml2js.Parser();

    parser.parseString(body, function(err, result) {
      var responseCode = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:ResponseCode'];

      if (responseCode !== 'NoError') {
        return callback(new Error(responseCode));
      }

      var item = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:Items']['t:Message'];

      var itemId = {
        id: item['t:ItemId']['@'].Id,
        changeKey: item['t:ItemId']['@'].ChangeKey
      };

      function handleMailbox(mailbox) {
        var mailboxes = [];

        if (!mailbox || !mailbox['t:Mailbox']) {
          return mailboxes;
        }
        mailbox = mailbox['t:Mailbox'];

        function getMailboxObj(mailboxItem) {
          return {
            name: mailboxItem['t:Name'],
            emailAddress: mailboxItem['t:EmailAddress']
          };
        }

        if (mailbox instanceof Array) {
          mailbox.forEach(function(m, idx) {
            mailboxes.push(getMailboxObj(m));
          });
        } else {
          mailboxes.push(getMailboxObj(mailbox));
        }

        return mailboxes;
      }

      var toRecipients = handleMailbox(item['t:ToRecipients']);
      var ccRecipients = handleMailbox(item['t:CcRecipients']);
      var from = handleMailbox(item['t:From']);

      var email = {
        id: itemId.id + '|' + itemId.changeKey,
        subject: item['t:Subject'],
        bodyType: item['t:Body']['@']['t:BodyType'],
        body: item['t:Body']['#'],
        size: item['t:Size'],
        dateTimeSent: item['t:DateTimeSent'],
        dateTimeCreated: item['t:DateTimeCreated'],
        toRecipients: toRecipients,
        ccRecipients: ccRecipients,
        from: from,
        isRead: (item['t:IsRead'] == 'true') ? true : false,
        meta: {
          itemId: itemId
        }
      };

      callback(null, email);
    });
  });
};

exports.getFolders = function(id, callback) {
  if (typeof(id) == 'function') {
    callback = id;
    id = 'inbox';
  }

  var soapRequest =
    '<tns:FindFolder Traversal="Shallow" xmlns:tns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
        '<tns:FolderShape>' +
          '<t:BaseShape>Default</t:BaseShape>' +
        '</tns:FolderShape>' +
        '<tns:ParentFolderIds>' +
          '<t:DistinguishedFolderId Id="inbox"></t:DistinguishedFolderId>' +
        '</tns:ParentFolderIds>' +
      '</tns:FindFolder>';

  exports.client.FindFolder(soapRequest, function(err, result) {
    if (err) {
      callback(err);
    }

    if (result.ResponseMessages.FindFolderResponseMessage.ResponseCode == 'NoError') {
      var rootFolder = result.ResponseMessages.FindFolderResponseMessage.RootFolder;

      rootFolder.Folders.Folder.forEach(function(folder) {
        // console.log(folder);
      });

      callback(null, {});
    }
  });
};

var fieldBuilder = function(fields) {
  return {
    buildFields: function() {
      return fields.reduce(function(memo, item) {
        return memo + '<t:FieldURI FieldURI="' + item + '" />';
      }, '');
    },

    processResult: function(item) {
      return fields.reduce(function(memo, field) {
        var propName = field.replace(/[a-z]+:/, '');
        var val = item['t:' + propName];
        if(val) {
          var camelCased = propName.charAt(0).toLowerCase() + propName.substr(1);
          memo[camelCased] = item['t:' + propName];
        }
        return memo;
      }, {});
    }
  };
};


var findItem = function(request, callback) {
  exports.client.FindItem(request, function(err, _, body) {
    if (err) {
      return callback(err);
    }
    var parser = new xml2js.Parser();
    parser.parseString(body, function(err, result) {
      if (err) {
        callback(err);
      } else {
        callback(null, result);
      }
    });
  });
};

var errorFactory = function(queryFn) {
  return function(event) {
    if(event.hasValue()) {
      var result = event.value();
      var responseCode = queryFn(result);
      if (responseCode !== 'NoError') {
        this.push(new Bacon.Error(responseCode));
        return;
      }
    }
    this.push(event);
  };
};

exports.getCalendarItems = function(startDate, endDate, callback) {

  var calendarFolderRequest =
    '<tns:GetFolder>' +
      '<tns:FolderShape>' +
        '<t:BaseShape>IdOnly</t:BaseShape>' +
      '</tns:FolderShape>' +
      '<tns:FolderIds>' +
      '<t:DistinguishedFolderId Id="calendar" />' +
      '</tns:FolderIds>' +
    '</tns:GetFolder>';

  var calendarFolderRequestErrorHandler = errorFactory(function(result) {
    return result['s:Body']['m:GetFolderResponse']['m:ResponseMessages']['m:GetFolderResponseMessage']['m:ResponseCode'];
  });

  var extractIdAndChangeKey = function(result) {
    var folderItem = result['s:Body']['m:GetFolderResponse']['m:ResponseMessages']['m:GetFolderResponseMessage']['m:Folders']['t:CalendarFolder']['t:FolderId'];
    return {
      folderId:folderItem['@'].Id,
      changeKey:folderItem['@'].ChangeKey
    };
  };

  // First request succeeded and received folderId and changeKey. Now get calendatrItems
  var calendarReq = fieldBuilder([
    'item:Subject',
    'calendar:Start',
    'calendar:End',
    'calendar:Duration',
    'calendar:Location',
    'calendar:Organizer'
  ]);

  var createItemsRequest = function(result) {
    return '<tns:FindItem Traversal="Shallow">' +
        '<tns:ItemShape>' +
          '<t:BaseShape>IdOnly</t:BaseShape>' +
          '<t:AdditionalProperties>' + calendarReq.buildFields() + '</t:AdditionalProperties>' +
        '</tns:ItemShape>' +
        '<tns:CalendarView StartDate="' + startDate + '" EndDate="' + endDate + '"/>' +
        '<tns:ParentFolderIds>' +
          '<t:FolderId Id="' + result.folderId + '" ChangeKey="' + result.changeKey + '" />' +
        '</tns:ParentFolderIds>' +
      '</tns:FindItem>';
  };

  var calendarItemRequestErrorHandler = errorFactory(function(result) {
    return result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:ResponseCode'];
  });

  var extractMailboxName = function(item) {
    return item['t:Mailbox']['t:Name'];
  };

  var extractAttendees = function(item) {
    var attendees = item['t:Attendee'];
    attendees = Array.isArray(attendees) ? attendees : [attendees];
    return attendees.map(extractMailboxName);
  };

  var createCalendarItem = function(result) {
    var rootFolder = result['s:Body']['m:FindItemResponse']['m:ResponseMessages']['m:FindItemResponseMessage']['m:RootFolder'];
    return rootFolder['t:Items']['t:CalendarItem'].map(function(item, idx) {
      var itemId = {
        id: item['t:ItemId']['@'].Id,
        changeKey: item['t:ItemId']['@'].ChangeKey
      };

      var calendarItem = calendarReq.processResult(item);
      calendarItem.id = itemId;
      calendarItem.organizer = extractMailboxName(calendarItem.organizer);
      return calendarItem;
    });
  };

  var itemReq = fieldBuilder([
    'item:Body',
    'calendar:RequiredAttendees',
    'calendar:OptionalAttendees',
  ]);

  var createItemRequest = function(result) {
    return '<tns:GetItem>' +
      '<tns:ItemShape>' +
        '<t:BaseShape>IdOnly</t:BaseShape>' +
        '<t:AdditionalProperties>' + itemReq.buildFields() + '</t:AdditionalProperties>' +
      '</tns:ItemShape>' +
      '<tns:ItemIds>' +
        '<t:ItemId Id="' + result.id.id + '" ChangeKey="' + result.id.changeKey + '"/>' +
      '</tns:ItemIds>' +
    '</tns:GetItem>';
  };

  var calendarItemDetailsErrorHandler = errorFactory(function(result) {
    return result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:ResponseCode'];
  });

  var createCalendarItemDetails = function(mergeTarget, result) {
    var itemResponse = result['s:Body']['m:GetItemResponse']['m:ResponseMessages']['m:GetItemResponseMessage']['m:Items']['t:CalendarItem'];
    var res = itemReq.processResult(itemResponse);
    for(var i in res) {
      if(res.hasOwnProperty(i)) {
        mergeTarget[i] = res[i];
      }
    }

    mergeTarget.requiredAttendees = mergeTarget.requiredAttendees ? extractAttendees(mergeTarget.requiredAttendees) : [];
    mergeTarget.optionalAttendees = mergeTarget.optionalAttendees ? extractAttendees(mergeTarget.optionalAttendees) : [];

    return mergeTarget;
  };

  var calendarFolderStream = Bacon.fromNodeCallback(findItem, calendarFolderRequest)
    .withHandler(calendarFolderRequestErrorHandler)
    .map(extractIdAndChangeKey)
    .map(createItemsRequest)
    .flatMap(function(calendarItemsRequest) {
      return Bacon.fromNodeCallback(findItem, calendarItemsRequest);
    })
    .withHandler(calendarItemRequestErrorHandler)
    .map(createCalendarItem)
    .flatMap(function(items) {
      return Bacon.fromArray(items)
        // Need to use a concurrency limit since EWS only allows 10 max simultaneous connections
        .flatMapWithConcurrencyLimit(5, function(item) {
          return Bacon.fromNodeCallback(findItem, createItemRequest(item))
            .withHandler(calendarItemDetailsErrorHandler)
            .map(function(result) {
              return createCalendarItemDetails(item, result);
            });
        })
        // Collect up all the annotated items into an array for returning to the client.
        .fold([], function (arr, meeting) {
          return arr.concat([meeting]);
        });
    });

  calendarFolderStream.onValue(callback.bind(null, null));
  calendarFolderStream.onError(callback);
};