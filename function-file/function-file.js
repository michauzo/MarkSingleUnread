/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {

};

// Add any ui-less function here
export function markAsUnread(event) {
    console.trace("MarkAsUnread: started");
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            var accessToken = result.value;

            console.trace("MarkAsUnread: Got token");
            // Use the access token
            markAsUnreadInternal(accessToken);
        } else {
            // Handle the error
            console.error(result.error);
        }
    });
}

function markAsUnreadInternal(accessToken) {
    // Get the item's REST ID
    console.trace("MarkAsUnread: internal started");
    var itemId = getItemRestId();

    // Construct the REST URL to the current item
    // Details for formatting the URL can be found at
    // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-a-message-rest
    var getMessageUrl = Office.context.mailbox.restUrl +
        '/v2.0/me/messages/' + itemId;

    $.ajax({
        url: getMessageUrl,
        method: 'PATCH',
        contentType: 'application/json',
        dataType: 'json',
        data: `{
    "IsRead": false
}`,
        headers: { 'Authorization': 'Bearer ' + accessToken }
    }).done(function (item) {
        // Message is passed in `item`
    }).fail(function (error) {
        // Handle error
        console.error(error);
    });
}

function getItemRestId() {
    console.trace("MarkAsUnread: Getting item id");
    if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
        // itemId is already REST-formatted
        return Office.context.mailbox.item.itemId;
    } else {
        // Convert to an item ID for API v2.0
        return Office.context.mailbox.convertToRestId(
            Office.context.mailbox.item.itemId,
            Office.MailboxEnums.RestVersion.v2_0
        );
    }
}