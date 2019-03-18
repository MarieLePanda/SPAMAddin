// The following string is a valid SOAP envelope and request for getting the properties of a mail item
function getItemDataSoap() {
    return        '<?xml version="1.0" encoding="utf-8"?>' +
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
        '      </ItemShape>' +
        '      <ItemIds>' +
        '        <t:ItemId Id="' + Office.context.mailbox.item.itemId + '"/>' +
        '      </ItemIds>' +
        '    </GetItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';
}

// The following string is a valid SOAP envelope and request for deleteing a mail item
function getForwardItemSoap(addressesSoap, bodyEmail, changeKey) {
    return '<?xml version="1.0" encoding="utf-8"?>' +
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
        '        <t:ForwardItem>' +
        '          <t:ToRecipients>' + addressesSoap + '</t:ToRecipients>' +
        '          <t:ReferenceItemId Id="' + Office.context.mailbox.item.itemId + '" ChangeKey="' + changeKey + '" />' +
        '          <t:NewBodyContent BodyType="Text">' + bodyEmail + '</t:NewBodyContent>' +
        '        </t:ForwardItem>' +
        '      </m:Items>' +
        '    </m:CreateItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';
}

// The following string is an email address come from API Config
function getEmailAddress() {
    var result;
    var xhr_object = new XMLHttpRequest();
    xhr_object.open("GET", "../api/Config/email", false);
    xhr_object.send();
    if (xhr_object.readyState == 4) {
        return JSON.parse(xhr_object.response);
    } else {
        return
        {
            Success: "Error"
            Data: "xhr_object.readyState is " + xhr_object.readyState
        };
    }
}

// The following string is a body email come from API Config
function getBodyEmail() {
    var result;
    var xhr_object = new XMLHttpRequest();
    xhr_object.open("GET", "../api/Config/bodyemail", false);
    xhr_object.send();
    if (xhr_object.readyState == 4) {
        return JSON.parse(xhr_object.response);
    } else {
        return
        {
            Success: "Error"
            Data: "xhr_object.readyState is " + xhr_object.readyState
        };
    }
}

// This function retrieves the current mail item, and the mailbox
// make an EWS request to get more properties of the item.
function FowardNow() {
    var soapToGetItemData = getItemDataSoap();

    Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, soapToGetItemDataCallback);
}

// This function is the callback for the makeEwsRequestAsync method
// checks for an error repsonse, but if all is OK it then parses the XML repsonse 
// to extract the ChangeKey attribute of the mailbox element.
function soapToGetItemDataCallback(asyncResult) {

    var parser;
    var xmlDoc;

    if (asyncResult.error != null) {
        statusUpdate("icon16", "the mail could not be forward to security team"
            + "error code is " + asyncResult.error.code);    }
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
        var changeKey = xmlDoc.getElementsByTagName("t:ItemId")[0].getAttribute("ChangeKey");
        var securityEmail;
        var bodyEmail;
        var response = getEmailAddress();

        if (response.Status == "Success") {
            securityEmail = response.Data;
        } else {
            statusUpdate("icon16", "the mail could not be forward to security team. "
                + "Error message : " + response.Data);
            return;
        }

        response = getBodyEmail();
        if (response.Status == "Success") {
            bodyEmail = response.Data;
        } else {
            statusUpdate("icon16", "the mail could not be forward to security team. "
                + "Error message : " + response.Data);
            return;
        }

        var addressesSoap = "<t:Mailbox><t:EmailAddress>" + securityEmail + "</t:EmailAddress></t:Mailbox>";
        var soapToForwardItem = getForwardItemSoap(addressesSoap, bodyEmail, changeKey);

        Office.context.mailbox.makeEwsRequestAsync(soapToForwardItem, soapToForwardItemCallback);

    }
}

// This function is the callback for the above makeEwsRequestAsync method
// In brief, it first checks for an error repsonse, but if all is OK
// it then parses the XML repsonse to extract the m:ResponseCode value.
function soapToForwardItemCallback(asyncResult) {

    var parser;
    var xmlDoc;

    if (asyncResult.error != null) {
        statusUpdate("icon16", "the mail could not be forward to security team"
            + "error code is " + asyncResult.error.code);
        return;
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

        var result = xmlDoc.getElementsByTagName("m:ResponseCode")[0].textContent;

        if (result == "NoError") {
            statusUpdate("icon16", "Email send to Security team, thanks for your action.");
        }
        else {
            statusUpdate("icon16", "the mail could not be forward to security team"
                + "error code is " + result);
            return;
        }
    }
}

function forwardStatus() {
    statusUpdate("icon16", "Email send to Security team, thanks for your action.");
}