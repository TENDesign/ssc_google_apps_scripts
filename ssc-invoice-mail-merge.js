const RECIPIENT_COL = "email";
const EMAIL_SENT_COL = "invoice_sent";
const DRIVE_FOLDER_ID = '1kQ-W0tMsCIwiPlzSUQ-X-mNz78MpAHVv'

// Creates the menu item "Mail Merge" for user to run scripts on drop-down.
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Send Invoices')
        .addItem('Mail Merge - Send Invoices', 'sendEmails')
        .addItem('Mail Merge - Send Invoice To Rows', 'sendEmailsToSelectedRows')
        .addToUi();
}

function sendEmailsToSelectedRows(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
    let selectedRows = Browser.inputBox('Send to rows comma seperated...',
        `Type or copy/paste a comma seperated list of rows to send to:`,
        Browser.Buttons.OK_CANCEL);
    if (selectedRows === "cancel" || selectedRows == '') return;
    selectedRows = selectedRows && [1, ...selectedRows.split(',').map(row => +row.trim())];
    sendEmails(subjectLine, sheet, selectedRows)
}

function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet(), selectedRows) {
    // option to skip browser prompt if you want to use this code in other projects
    if (!subjectLine) {
        subjectLine = Browser.inputBox('Mail Merge',
            `Type or copy/paste the subject line of the Gmail draft message you would like to mail merge with:`,
            Browser.Buttons.OK_CANCEL);

        if (subjectLine === "cancel" || subjectLine == '') return;
    }

    

    let data = [];

    // Gets the data from the passed sheet
    if (!selectedRows) data = sheet.getDataRange().getDisplayValues();
    if (selectedRows && selectedRows.length > 0) 
    selectedRows.forEach((row, index) => {
        const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
        data = [...data, ...rowData]
    })

    // Assumes row 1 contains our column headings
    const heads = data.shift();

    // Gets the index of the column named 'Email Status' (Assumes header names are unique)
    const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

    // Converts array into an object array
    const obj = data.map((row, i) => {
        return (
            heads.reduce((acc, curr, index) => 
                (
                    acc[curr] = row[index] || '',
                    acc
                ), {}
            )
        )
    });

    // Creates an array to record sent emails
    let out = [];

    // Gets the draft Gmail message to use as a template
    const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

    // Loops through all the rows of data
    obj.forEach((row) => {
        // Only sends emails if email_sent cell is blank and not hidden by a filter
        try {
            const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
            console.log(obj)

            const blob = Utilities.newBlob(msgObj.html, MimeType.HTML, `${obj[0].last},${obj[0].first} - ${new Date().toDateString()} `).getAs(MimeType.PDF)
            DriveApp.getFolderById(DRIVE_FOLDER_ID).createFile(blob);


            // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
            // Uncomment advanced parameters as needed (see docs for limitations)
            GmailApp.sendEmail(row[RECIPIENT_COL], `Sandusky Sailing Club | Sadler Basin: ${msgObj.subject}`, msgObj.text, {
                htmlBody: msgObj.html,
                from: 'sanduskysailingclub@tendesign.us',
                name: 'Sandusky Sailing Club',
                attachments: [blob],
                inlineImages: emailTemplate.inlineImages
                // bcc: cc: replyTo: noReply etc...,
            });
            // Edits cell to record email sent date
            out = [[new Date().toDateString()]];
        } catch (error) {
            // modify cell to record error
            out = [[error.message]];
        }
        // Updates the sheet with new data
        sheet.getRange(+row['row_id'], emailSentColIdx + 1, out.length).setValues(out);
    });


    function getGmailTemplateFromDrafts_(subject_line) {
        try {
            // get drafts
            const drafts = GmailApp.getDrafts();
            // filter the drafts that match subject line
            const draft = drafts.filter(subjectFilter_(subject_line))[0];
            // get the message object
            const msg = draft.getMessage();

            // Handles inline images and attachments so they can be included in the merge
            // Gets all attachments and inline image attachments
            const allInlineImages = draft.getMessage().getAttachments({ includeInlineImages: true, includeAttachments: false });
            const attachments = draft.getMessage().getAttachments({ includeInlineImages: false });
            const htmlBody = msg.getBody();

            // Creates an inline image object with the image name as key 
            // (can't rely on image index as array based on insert order)
            const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

            //Regexp searches for all img string positions with cid
            const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
            const matches = [...htmlBody.matchAll(imgexp)];

            //Initiates the allInlineImages object
            const inlineImagesObj = {};
            // built an inlineImagesObj from inline image matches
            matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

            return {
                message: { subject: subject_line, text: msg.getPlainBody(), html: htmlBody },
                attachments: attachments, inlineImages: inlineImagesObj
            };
        } catch (e) {
            throw new Error("Oops - can't find Gmail draft");
        }

        function subjectFilter_(subject_line) {
            return function (element) {
                if (element.getMessage().getSubject() === subject_line) {
                    return element;
                }
            }
        }
    }

    function fillInTemplateFromObject_(template, data) {
        // We have two templates one for plain text and the html body
        // Stringifing the object means we can do a global replace
        let template_string = JSON.stringify(template);

        // Token replacement
        template_string = template_string.replace(/{{[^{}]+}}/g, key => {
            const keyString = key.toString();
            const replaced =  escapeData_(data[keyString.replace(/[{}]+/g, "")] || "");
            return replaced
        });
        return JSON.parse(template_string);
    }

    function escapeData_(str) {
        return str
            // .replace(/[\\]/g, '\\\\')
            // .replace(/[\"]/g, '\\\"')
            // .replace(/[\/]/g, '\\/')
            // .replace(/[\b]/g, '\\b')
            // .replace(/[\f]/g, '\\f')
            // .replace(/[\n]/g, '\\n')
            // .replace(/[\r]/g, '\\r')
            // .replace(/[\t]/g, '\\t');
    };
}
