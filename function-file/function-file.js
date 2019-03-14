/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import $ from "jquery";

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
    window.markAsUnread = markAsUnread;
    window.flag = flag;
};

// Add any ui-less function here
function markAsUnread(event) {
    callWithToken(markAsUnreadInternal, event);
}

function flag(event) {
    callWithToken(flagInternal, event);
}

function callWithToken(func, event) {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            const accessToken = result.value;

            // Use the access token
            func(accessToken, () => event.completed());
        } else {
            // Handle the error
            console.error(result.error);
        }
    });
}

function markAsUnreadInternal(accessToken, callback) {
    const data = `{
    "IsRead": false
}`

    updateMessage(data, accessToken, callback);
}

function flagInternal(accessToken, callback) {
    const data = `{
    "Flag": {
        "FlagStatus": "Flagged"
    }
}`
    updateMessage(data, accessToken, callback);
}

function updateMessage(data, accessToken, callback) {
    // Get the item's REST ID
    const itemId = getItemRestId();

    // Construct the REST URL to the current item
    // Details for formatting the URL can be found at
    // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-a-message-rest
    const getMessageUrl = Office.context.mailbox.restUrl +
        '/v2.0/me/messages/' + itemId;

    $.ajax({
        url: getMessageUrl,
        method: 'PATCH',
        contentType: 'application/json',
        dataType: 'json',
        data,
        headers: { 'Authorization': 'Bearer ' + accessToken }
    }).done(function (item) {
        // Message is passed in `item`
        callback();
    }).fail(function (error) {
        // Handle error
        console.error(error);
        callback();
    });
}

function getItemRestId() {
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