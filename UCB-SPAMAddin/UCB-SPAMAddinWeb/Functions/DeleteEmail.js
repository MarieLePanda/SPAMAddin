// The following string is a valid SOAP envelope and request for getting the properties of a mail item
function getItemDataSoap() {
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
function getDeleteItemSoap(changeKey) {
    return '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '    <soap:Header>' +
        '        <t:RequestServerVersion Version="Exchange2013" />' +
        '    </soap:Header>' +
        '    <soap:Body>' +
        '        <m:MarkAsJunk IsJunk="true" MoveItem="true">' +
        '            <m:ItemIds>' +
        '                <t:ItemId Id="' + Office.context.mailbox.item.itemId + '" ChangeKey="' + changeKey + '" />' +
        '            </m:ItemIds>' +
        '        </m:MarkAsJunk>' +
        '   </soap:Body>' +
        '</soap:Envelope>';
}

// This function retrieves the current mail item, and the mailbox
// make an EWS request to get more properties of the item.
function DeleteNow() {

    var item = Office.context.mailbox.item;
    var soapToGetItemData = getItemDataSoap();

    Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, soapToGetItemDataCallbackForDelete);
}

// This function is the callback for the makeEwsRequestAsync method
// checks for an error repsonse, but if all is OK it then parses the XML repsonse 
// to extract the ChangeKey attribute of the mailbox element.
function soapToGetItemDataCallbackForDelete(asyncResult) {

    var parser;
    var xmlDoc;

    if (asyncResult.error != null) {
        statusUpdate("icon16", "the mail could not be delete "
            + "error code is " + asyncResult.error.code);
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

        var changeKey = xmlDoc.getElementsByTagName("t:ItemId")[0].getAttribute("ChangeKey");
        var soapToDeleteItem = getDeleteItemSoap(changeKey);
            
        Office.context.mailbox.makeEwsRequestAsync(soapToDeleteItem, soapToDeleteItemCallback);
    }
}

// This function is the callback for the above makeEwsRequestAsync method
// In brief, it first checks for an error repsonse, but if all is OK
// it then parses the XML repsonse to extract the m:ResponseCode value.
function soapToDeleteItemCallback(asyncResult) {

    var parser;
    var xmlDoc;

    if (asyncResult.error != null) {
        statusUpdate("icon16", "the mail could not be delete "
            + "error code is " + asyncResult.error.code);
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
            statusUpdate("icon16", "Email delete , thanks for your action.");
        }
        else {
            statusUpdate("icon16", "the mail could not be delete "
                + "error code is " + result);
            return;
        }
    }
}

function deleteStatus() {
    statusUpdate("icon16", "Email delete , thanks for your action.");
}