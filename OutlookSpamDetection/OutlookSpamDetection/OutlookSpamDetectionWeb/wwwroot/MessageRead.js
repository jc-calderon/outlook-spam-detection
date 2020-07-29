(function () {
    "use strict";

    var messageBanner;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();
            loadProps();
        });
    };

    // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
    function buildAttachmentsString(attachments) {
        if (attachments && attachments.length > 0) {
            var returnString = "";

            for (var i = 0; i < attachments.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + attachments[i].name;
            }

            return returnString;
        }

        return "None";
    }

    // Format an EmailAddressDetails object as
    // GivenName Surname <emailaddress>
    function buildEmailAddressString(address) {
        return address.displayName + " &lt;" + address.emailAddress + "&gt;";
    }

    // Take an array of EmailAddressDetails objects and
    // build a list of formatted strings, separated by a line-break
    function buildEmailAddressesString(addresses) {
        if (addresses && addresses.length > 0) {
            var returnString = "";

            for (var i = 0; i < addresses.length; i++) {
                if (i > 0) {
                    returnString = returnString + "<br/>";
                }
                returnString = returnString + buildEmailAddressString(addresses[i]);
            }

            return returnString;
        }

        return "None";
    }

    // Load properties from the Item base object, then load the
    // message-specific properties.
    function loadProps() {
        var item = Office.context.mailbox.item;
        var emailInfo = {
            from: {
                displayName: item.from.displayName,
                emailAddress: item.from.emailAddress
            },
            to: {
                displayName: item.to[0].displayName,
                emailAddress: item.to[0].emailAddress
            },
            subject: item.subject,
            bodyText: ""
        }

        $('#message-props').show();
        $('#attachments').html(buildAttachmentsString(item.attachments));
        $('#cc').html(buildEmailAddressesString(item.cc));
        $('#from').html(buildEmailAddressString(item.from));
        $('#sender').html(buildEmailAddressString(item.sender));
        $('#subject').text(item.subject);
        $('#to').html(buildEmailAddressesString(item.to));

        item.body.getAsync('text', function (result) {
            if (result.status === 'succeeded') {
                emailInfo.bodyText = result.value;
                $('#bodyText').text(emailInfo.bodyText);

                requestPOST('/api/email', {}, emailInfo, onSuccess, onError);
            }
        });

        function onSuccess(success) {
            console.log(JSON.stringify(success))
            if (success.isSpam) {
                $('#notificationBody').addClass('red-text');
                showNotification("Outlook Spam Detection", "This Email is SPAM");
            } else {
                showNotification("Outlook Spam Detection", "This  Email is not spam");
            }
        }

        function onError(error) {
        }
    }

    function requestPOST(url, headers, data, onSuccess, onError) {
        headers['Accept'] = 'application/json';
        headers['Content-Type'] = 'application/json';

        $.ajax({
            url: url,
            headers: headers,
            method: "POST",
            data: JSON.stringify(data)
        }).done(function (response) {
            if (response) {
                onSuccess(response);
            } else {
                onSuccess();
            }
        }).fail(function (error) {
            if (onError) {
                onError(error);
            } else {
                onError();
            }
        });
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();