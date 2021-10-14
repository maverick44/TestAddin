var mailboxItem

Office.initialize = function () {

   
}

function OpenNewEMailDialog() {
    Office.context.mailbox.displayNewMessageForm(
        {
            //toRecipients: Office.context.mailbox.item.to, // Copies the To line from current item
            toRecipients: ["md.usman@gmail.com"], // Copies the To line from current item
            //ccRecipients: ["md.usman@gmail.com"],
            subject: "Phish Email",
            htmlBody: 'Hello <b>World</b>!<br/></i>',
            attachments:
                [
                    { type: "item", itemId: Office.context.mailbox.item.itemId, name: "PhishEmail.msg" }
                ],
            options: { asyncContext: null },
            callback: function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showMessage("Action failed with error: " + asyncResult.error.message);
                }
            }
        });
}

function sendEmailNow() {
    
    //Office.context.mailbox.item.notificationMessages.replaceAsync("status1", {
    //    type: "informationalMessage",
    //    icon: "icon16",
    //    message: "Kicking in now",
    //    persistent: false
    //});

    //OpenNewEMailDialog();

    sendEmailAttach();
    
    //OpenNewEMailDialog();
   
    //Office.context.mailbox.item.notificationMessages.removeAsync("status");
}

function deleteSelectedEmail() {
    var item = Office.context.mailbox.item;
    var item_id = item.itemId;
    var mailbox = Office.context.mailbox;

    var reqDelete1 = '<?xml version="1.0" encoding="utf-8"?> ' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
        'xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
        'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
        'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
        '<soap:Header> ' +
        '<t:RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" /> ' +
        '</soap:Header> ' +
        '<soap:Body> ' +
        '<MoveItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
        'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
        '<ToFolderId> ' +
        '<t:DistinguishedFolderId Id="deleteditems"/> ' +
        '</ToFolderId> ' +
        '<ItemIds> ' +
        '<t:ItemId Id="' + item_id + '"/> ' +
        '</ItemIds> ' +
        '</MoveItem> ' +
        '</soap:Body> ' +
        '</soap:Envelope> ';

    // The makeEwsRequestAsync method accepts a string of SOAP and a callback function
    mailbox.makeEwsRequestAsync(reqDelete1, soapDeleteResponse);
}

function soapDeleteResponse(asyncResult) {
    var parser;
    var xmlDoc;

    if (asyncResult.error != null) {

        console.log(asyncResult.error.message+ " 123");
        //app.showNotification("EWS Status", asyncResult.error.message);        
    }
    else {
        var response = asyncResult.value;
        if (window.DOMParser) {
            parser = new DOMParser();
            xmlDoc = parser.parseFromString(response, "text/xml");
        }
        else // Older Versions of Internet Explorer
        {
            xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
            xmlDoc.async = false;
            xmlDoc.loadXML(response);
        }

        // Get the required response, and if it's NoError then all has succeeded, so tell the user.
        // Otherwise, tell them what the problem was. (E.G. Recipient email addresses might have been
        // entered incorrectly --- try it and see for yourself what happens!!)
        var result = xmlDoc.getElementsByTagName("m:ResponseCode")[0].textContent;
        if (result == "NoError") {
            //    app.showNotification("EWS Status", "Success!");
        }
        else {
            console.log(result + " 456");
            //    app.showNotification("EWS Status", "The following error code was recieved: " + result);
        }
    }
}

function sendEmailAttach() {
    var item = Office.context.mailbox.item;
    var item_id = item.itemId;
    var mailbox = Office.context.mailbox;

    // The following string is a valid SOAP envelope and request for getting the properties
    // of a mail item. Note that we use the item_id value (which we obtained above) to specify the item
    // we are interested in.
    var soapToGetItemData = '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <GetItem' +
        '                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '      <ItemShape>' +
        '        <t:BaseShape>IdOnly</t:BaseShape>' +

        '        <t:AdditionalProperties>' +
        '        <t:FieldURI FieldURI="item:MimeContent"/>' +
        '        <t:FieldURI FieldURI="item:Subject"/>' +
        '        </t:AdditionalProperties>' +

        '      </ItemShape>' +
        '      <ItemIds>' +
        '        <t:ItemId Id="' + item_id + '"/>' +
        '      </ItemIds>' +
        '    </GetItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';

    // The makeEwsRequestAsync method accepts a string of SOAP and a callback function
    mailbox.makeEwsRequestAsync(soapToGetItemData, soapToGetItemDataCallback);
}

// This function is the callback for the makeEwsRequestAsync method
// In brief, it first checks for an error repsonse, but if all is OK
// it then parses the XML repsonse to extract the ChangeKey attribute of the 
// t:ItemId element.
function soapToGetItemDataCallback(asyncResult) {
    var parser;
    var xmlDoc;

    if (asyncResult.error != null)
    {
        Office.context.mailbox.item.notificationMessages.replaceAsync("status1", {
            type: "informationalMessage",
            icon: "icon16",
            message: "Could not forward the email.",
            persistent: false
        });
    }
    else
    {
        var response = asyncResult.value;
        if (window.DOMParser) {
            var parser = new DOMParser();
            xmlDoc = parser.parseFromString(response, "text/xml");
        }
        else // Older Versions of Internet Explorer
        {
            xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
            xmlDoc.async = false;
            xmlDoc.loadXML(response);
        }

        var emailSenderName1 = Office.context.mailbox.item.from.displayName;
        var emailSender1 = Office.context.mailbox.item.from.emailAddress;
        var mimeContentID1 = xmlDoc.getElementsByTagName("t:MimeContent")[0].textContent;
        var emailSubject1 = xmlDoc.getElementsByTagName("t:Subject")[0].textContent;
        //app.showNotification(emailSubject1, mimeContentID1);   

        var toAddress1 = "<t:Mailbox><t:EmailAddress>phishing@develsecurity.com</t:EmailAddress></t:Mailbox>"

        // The following string is a valid SOAP envelope and request for forwarding
        // a mail item. Note that we use the item_id value (which we obtained in the click event handler)
        // to specify the item we are interested in,
        // along with its ChangeKey value that we have just determined near the top of this function.
        // We also provide the XML fragment that we built in the loop above to specify the recipient addresses,
        // and the comment that the user may have provided in the Comment: text box

        var newTestSendEmail = '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
            '  </soap:Header>' +
            '  <soap:Body>' +
            '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
            '      <m:Items>' +
            '       <t:Message>' +
            ' <t:Subject>Phishing Email from Eyephish Add-in</t:Subject>' +
            ' <t:Body BodyType="Text">Sender Name: ' + emailSenderName1 +'\rSender Email: ' + emailSender1 + '\rEmail Subject: ' + emailSubject1 + '</t:Body> ' +
            ' <t:ToRecipients>' + toAddress1 + '</t:ToRecipients>' +

            ' <t:Attachments>' +
            ' <t:ItemAttachment>' +
            ' <t:Name>' + emailSubject1 + '</t:Name>' +
            ' <t:IsInline>false</t:IsInline>' +
            ' <t:Message>' +
            ' <t:MimeContent CharacterSet="UTF-8">' + mimeContentID1 + '</t:MimeContent>' +
            ' </t:Message>' +
            ' </t:ItemAttachment>' +
            ' </t:Attachments>' +

            '      </t:Message>' +
            '      </m:Items>' +
            '    </m:CreateItem>' +
            '  </soap:Body>' +
            '</soap:Envelope>';

        // As before, the makeEwsRequestAsync method accepts a string of SOAP and a callback function.
        // The only difference this time is that the body of the SOAP message requests that the item
        // be forwarded (rather than retrieved as in the previous method call)
        Office.context.mailbox.makeEwsRequestAsync(newTestSendEmail, soapToForwardItemCallback);
    }
}

// This function is the callback for the above makeEwsRequestAsync method
// In brief, it first checks for an error repsonse, but if all is OK
// it then parses the XML repsonse to extract the m:ResponseCode value.
function soapToForwardItemCallback(asyncResult) {
    var parser;
    var xmlDoc;

    if (asyncResult.error != null)
    {
        Office.context.mailbox.item.notificationMessages.replaceAsync("status1", {
            type: "informationalMessage",
            icon: "icon16",
            message: "Could not forward the email.",
            persistent: false
        });
        //app.showNotification("EWS Status", asyncResult.error.message);        
    }
    else {
        var response = asyncResult.value;
        if (window.DOMParser) {
            parser = new DOMParser();
            xmlDoc = parser.parseFromString(response, "text/xml");
        }
        else // Older Versions of Internet Explorer
        {
            xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
            xmlDoc.async = false;
            xmlDoc.loadXML(response);
        }

        // Get the required response, and if it's NoError then all has succeeded, so tell the user.
        // Otherwise, tell them what the problem was. (E.G. Recipient email addresses might have been
        // entered incorrectly --- try it and see for yourself what happens!!)
        var result = xmlDoc.getElementsByTagName("m:ResponseCode")[0].textContent;
        if (result == "NoError") {
            Office.context.mailbox.item.notificationMessages.replaceAsync("status1", {
                type: "informationalMessage",
                icon: "icon16",
                message: "Email has been forwarded as SPAM.",
                persistent: false
            });

            deleteSelectedEmail();
        }
        else {
            Office.context.mailbox.item.notificationMessages.replaceAsync("status1", {
                type: "informationalMessage",
                icon: "icon16",
                message: "Could not forward the email.",
                persistent: false
            });
            //    app.showNotification("EWS Status", "The following error code was recieved: " + result);
        }

       
    }
}
