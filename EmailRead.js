var DialogRead;

(function () {
    "use strict";

    var GEmailSenderName="";
    var GEmailSender="";
    var GEmailSubject="";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            ReadEmailSubject();
            //Office.context.mailbox.item.body.getAsync("html", processHtmlBody);
        });
    };

    function processHtmlBody(asyncResult)
    {
        var htmlParser = new DOMParser().parseFromString(asyncResult.value, "text/html");
        var links = htmlParser.getElementsByTagName("a");
        var phishyLinkCount = 0;
        var arrLinks = [];
        var str1 = "";
        var str2 = "";

        $.each(
            links,
            function (i, v)
            {
                var regExp = new RegExp('/+$');
                var vInnerText = v.innerText.toLowerCase().trim().replace(regExp, "");
                var hrefText = v.href.toLowerCase().trim().replace(regExp, "");;
                //var linkIsPhishy = ((vInnerText.search("http") == 0) && vInnerText != hrefText);

                arrLinks.push(hrefText);

                phishyLinkCount++;
                $("#links-table").append("<div class='ms-Table-row ms-font-xs ms-font-color-white'>" +
                    "<span class='ms-Table-cell phishy-link'>" + hrefText + "</span>" +
                    "</div>");
            }
        );

        $('#result').append("Number of Links Found: " + (phishyLinkCount));

        // 1st API Call

        jQuery.post("https://epapi.develsystems.com/SearchPhishingURL/", arrLinks, function (dataSender, statusSender) {
            var apiResultA = JSON.stringify(dataSender);
            var apiResultJsonA = jQuery.parseJSON(apiResultA);
            localStorage.setItem("key1", apiResultJsonA.Status);
        });
        str1 = localStorage.getItem("key1");

        // 2nd API Call

        var emailContents = "{\"Recipient_email\": \"" + GEmailSender + "\",\"email_subject\": \"" + GEmailSubject + "\",\"sender_email\": \"" + GEmailSenderName +"\"}";

        jQuery.post("https://epapi.develsystems.com/SearchPhishingSender/", emailContents, function (dataEmail, statusEmail) {
            var apiResultB = JSON.stringify(dataEmail);
            var apiResultJsonB = jQuery.parseJSON(apiResultB);
            localStorage.setItem("key2", apiResultJsonB.Status)
        });
        str2 = localStorage.getItem("key2");

        if (str1 == "1" || str2 == "1") {
            //sendEmailAttachRead();
        }

        Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog.html?" + str1 + str2,
            { height: 20, width: 30, displayInIframe: true }, dialogCallbackRead);
    }

    function dialogCallbackRead(asyncResult)
    {
        if (asyncResult.status == "failed")
        {
            switch (asyncResult.error.code)
            {
                case 12004:
                    showNotification("Domain is not trusted");
                    break;
                case 12005:
                    showNotification("HTTPS is required");
                    break;
                case 12007:
                    showNotification("A dialog is already opened.");
                    break;
                default:
                    showNotification(asyncResult.error.message);
                    break;
            }
        }
        else
        {
            DialogRead = asyncResult.value;
            /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
            DialogRead.addEventHandler(Office.EventType.DialogMessageReceived, messageHandlerRead);
            
            /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
            DialogRead.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
        }
    }


    function messageHandlerRead(arg) {
        DialogRead.close();
        var dialogRes = arg.message;

        if (dialogRes != "00") {
            sendEmailAttachRead();
        }
    }

    function sendEmailAttachRead() {
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
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types"/>' +
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
        mailbox.makeEwsRequestAsync(soapToGetItemData, soapToGetItemDataCallbackRead);
    }

    function soapToGetItemDataCallbackRead(asyncResult) {
        var parser;
        var xmlDoc;

        if (asyncResult.error != null) {
            Office.context.mailbox.item.notificationMessages.replaceAsync("status1", {
                type: "informationalMessage",
                icon: "icon16",
                message: "Could not forward the email.",
                persistent: false
            });
        }
        else {
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
            
            var toAddress1 = "<t:Mailbox><t:EmailAddress>phishing@develsecurity.com</t:EmailAddress></t:Mailbox>"

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
                ' <t:Body BodyType="Text">Sender Name: ' + emailSenderName1 + '\rSender Email: ' + emailSender1 + '\rEmail Subject: ' + emailSubject1 + '</t:Body> ' +
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
            Office.context.mailbox.makeEwsRequestAsync(newTestSendEmail, soapToForwardItemCallbackRead);
        }
    }

    function soapToForwardItemCallbackRead(asyncResult) {
        var parser;
        var xmlDoc;

        if (asyncResult.error != null) {
            Office.context.mailbox.item.notificationMessages.replaceAsync("status1", {
                type: "informationalMessage",
                icon: "icon16",
                message: "Could not forward the email.",
                persistent: false
            });
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
            if (result == "NoError")
            {
                deleteSelectedEmailRead();
            }
            else
            {
                Office.context.mailbox.item.notificationMessages.replaceAsync("status1",
                    {
                    type: "informationalMessage",
                    icon: "icon16",
                    message: "Could not forward the email.",
                    persistent: false
                });
            }
        }
    }

    function deleteSelectedEmailRead() {
        var item = Office.context.mailbox.item;
        var item_id = item.itemId;
        var mailbox = Office.context.mailbox;

        var reqDelete1 = '<?xml version="1.0" encoding="utf-8"?> ' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
            'xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
            'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ' +
            'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"> ' +
            '<soap:Header> ' +
            '<t:RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types"/> ' +
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

        mailbox.makeEwsRequestAsync(reqDelete1, soapDeleteResponseRead);
    }

    function soapDeleteResponseRead(asyncResult) {
        var parser;
        var xmlDoc;

        if (asyncResult.error != null)
        {
        }
        else
        {
            var response = asyncResult.value;
            if (window.DOMParser)
            {
                parser = new DOMParser();
                xmlDoc = parser.parseFromString(response, "text/xml");
            }
            else // Older Versions of Internet Explorer
            {
                xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                xmlDoc.async = false;
                xmlDoc.loadXML(response);
            }

            var result = xmlDoc.getElementsByTagName("m:ResponseCode")[0].textContent;
            if (result == "NoError")
            {
            }
            else
            {
            }
        }
    }

    function ReadEmailSubject()
    {
        var item = Office.context.mailbox.item;
        var item_id = item.itemId;
        var mailbox = Office.context.mailbox;

        var soapToGetItemData = '<?xml version="1.0" encoding="utf-8"?>' +
            '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
            '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
            '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
            '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
            '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
            '  <soap:Header>' +
            '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types"/>' +
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

        mailbox.makeEwsRequestAsync(soapToGetItemData, soapReadSubject);
    }

    function soapReadSubject(asyncResult)
    {
        var parser;
        var xmlDoc;

        if (asyncResult.error != null) {

        }
        else
        {
            var response = asyncResult.value;
            if (window.DOMParser)
            {
                var parser = new DOMParser();
                xmlDoc = parser.parseFromString(response, "text/xml");
            }
            else // Older Versions of Internet Explorer
            {
                xmlDoc = new ActiveXObject("Microsoft.XMLDOM");
                xmlDoc.async = false;
                xmlDoc.loadXML(response);
            }

            GEmailSenderName = Office.context.mailbox.item.from.displayName;
            GEmailSender = Office.context.mailbox.item.from.emailAddress;
            GEmailSubject = xmlDoc.getElementsByTagName("t:Subject")[0].textContent;

            Office.context.mailbox.item.body.getAsync("html", processHtmlBody);
        }
    }

})();

