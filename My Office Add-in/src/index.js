'use strict';

(function () {

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      loadItemProps(Office.context.mailbox.item);
      $('#test-me').click(() => debug(Office.context));
    });
  };

  function debug(context) {
    console.log("Hello World!");
    getRestToken(context.mailbox)
      .then(token => getTransportHeaders(context, token))
      .then(item => {
        console.log(item)})
  }

  function getRestToken(mailbox) {
    return new Promise((resolve, reject) => {
      mailbox.getCallbackTokenAsync({isRest: true}, result => {
        if (result.status === "succeeded") {
          resolve(result.value);
        } else {
          reject(result.status);
        }
      })
    })
  }

  function getItemRestId(context) {
    if (context.mailbox.diagnostics.hostName === "OutlookIOS") {
      return context.mailbox.item.itemId;
    } else {
      return context.mailbox.convertToRestId(
        context.mailbox.item.itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
    }
  }

  function getTransportHeaders(context, accessToken) {
    var itemId = getItemRestId(context);
    var getMessageUrl = context.mailbox.restUrl
      + '/v2.0/me/messages/' + itemId;
    return $.ajax({
      url: getMessageUrl,
      dataType: 'json',
      headers: {'Authorization': "Bearer " + accessToken}
    })
  }

  function loadItemProps(item) {
    // Get the table body element
    var tbody = $('.prop-table');

    // Add a row to the table for each message property
    tbody.append(makeTableRow("Id", item.itemId));
    tbody.append(makeTableRow("Subject", item.subject));
    tbody.append(makeTableRow("Message Id", item.internetMessageId));
    tbody.append(makeTableRow("From", item.from.displayName + " &lt;" +
      item.from.emailAddress + "&gt;"));
    // item.bcc.getAsync(bccRecipients => {
    //   const bccEmails = bccRecipients.value.map(r => r.emailAddress);
    //   tbody.append(makeTableRow("BCC", bccEmails));
    // })
    // item.loadCustomPropertiesAsync(x => {
    //   const ready = readline();
    //   const customProps = x.value;
    //   console.log(x);
    //   console.log(customProps.get("PR_TRANSPORT_MESSAGE_HEADERS"));
    // });
    /*
    var keys = [];
    for (var key in item) {
      tbody.append(makeTableRow(key, item[key]));
    }
    //*/
  }

  function makeTableRow(name, value) {
    return $("<tr><td><strong>" + name + 
      "</strong></td><td class=\"prop-val\"><code>" +
      value + "</code></td></tr>");
  }

})();