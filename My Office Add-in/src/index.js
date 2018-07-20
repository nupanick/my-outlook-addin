'use strict';

(function () {

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      loadItemProps(Office.context.mailbox.item);
    });
  };

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