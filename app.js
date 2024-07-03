var express = require('express'),
    app = express(),
    bodyParser = require('body-parser'),
    officeGen = require('officegen'),
    nodemailer = require('nodemailer'),
    fs = require('fs'),
    path = require('path');

const { google } = require('googleapis');
const MailComposer = require('nodemailer/lib/mail-composer');
const credentials = require('./credentials.json');
const tokens = require('./token.json');

const getGmailService = () => {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);
  oAuth2Client.setCredentials(tokens);
  const gmail = google.gmail({ version: 'v1', auth: oAuth2Client });
  return gmail;
};

const encodeMessage = (message) => {
  return Buffer.from(message).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
};

const createMail = async (options) => {
  const mailComposer = new MailComposer(options);
  const message = await mailComposer.compile().build();
  return encodeMessage(message);
};

const sendMail = async (options) => {
  const gmail = getGmailService();
  const rawMessage = await createMail(options);
  const { data: { id } = {} } = await gmail.users.messages.send({
    userId: 'chenweinberg@gmail.com',
    resource: {
      raw: rawMessage,
    },
  });
  return id;
};

const mail = async(email, service)  => {
    var mailOptions = {
        from: "Neshama Therapy <chenweinberg@gmail.com>",
        to: email,
        subject: "Your " + service + " Receipt For Today's Treatment ",
        text: "Thank You For Choosing Neshama Therapy, Your Receipt Is Attached With This Email. ",
        textEncoding: 'base64',
        attachments: [
            {
                filename: 'your-receipt.docx',
                path: __dirname + '/tmp/receipt.docx'
            }
        ]
    }
    var mailOptions2 = {
        from: "Neshama Therapy <chenweinberg@gmail.com>",
        to: "chenweinberg@gmail.com",
        subject: "Your " + service + " Receipt For Today's Treatment ",
        text: "Thank You For Choosing Neshama Therapy, Your Receipt Is Attached With This Email. ",
        textEncoding: 'base64',
        attachments: [
            {
                filename: 'your-receipt.docx',
                path: __dirname + '/tmp/receipt.docx'
            }
        ]
    }
    const messageId = await sendMail(mailOptions);
    const messageId2 = await sendMail(mailOptions2)
    return messageId;
}

app.set('views', __dirname + '/views');
app.set('view engine', 'ejs');
app.use(express.static(__dirname + '/public'));
app.use(bodyParser.urlencoded({ extended: true }));


app.get('/', function (req, res) {
    res.render('home');
});


app.post('/acu-receipt', function (req, res) {
    var docx = officeGen({
        'type': 'docx',
        'orientation': 'portrait',
        'subject': 'Neshama Therapy Receipt',
        'description': 'Neshama Therapy Receipt'
    });
    docx.on('error', function (err) {
        res.redirect('/');
    });
    let address = '55 Maitland St, Suite 1502, Toronto, ON M4Y1C9';
    if (req.body.new_address) {
        address = req.body.new_address;
    };
    var pObj = docx.createP();
    pObj.options.align = 'left';
    pObj.addImage(path.resolve(__dirname, 'neshamalogo.png'));
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addText(address, { bold: true });
    pObj.addLineBreak();
    pObj.addText('Chen Weinberg, R. Ac, RMT', { bold: true });
    pObj.addLineBreak();
    pObj.addText('Acupuncture Registration Number: 2295');
    pObj.addLineBreak();
    pObj.addText('College of Traditional Chinese Medicine and Acupuncturists of Ontario');
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addHorizontalLine();
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addText('ACUPUNCTURE RECEIPT', { color: '07421e', bold: true, font_size: 14 });
    pObj.addLineBreak();
    pObj.addText('Patient Name: ' + req.body.patient_name, { color: '06a2db', bold: true, font_size: 18 });
    pObj.addLineBreak();
    pObj.addText('Date: ' + req.body.date, { color: '204903', bold: true });
    pObj.addLineBreak();
    pObj.addText('Duration of Treatment: ' + req.body.service_length, { color: '204903', bold: true });
    pObj.addLineBreak();
    pObj.addText('Description Of Service: ' + req.body.service_desc, { color: '204903', bold: true });
    pObj.addLineBreak();
    pObj.addText('Total Amount: ' + req.body.amount, { color: '204903', bold: true });
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addText('Sincerely,');
    pObj.addLineBreak();
    pObj.addText('Chen Weinberg, R. Ac, RMT', { color: '10350f', bold: true });
    pObj.addLineBreak();
    pObj.addImage(path.resolve(__dirname, 'signature.JPG'));

    var out = fs.createWriteStream(__dirname + '/tmp/receipt.docx');
    out.on('error', function (err) {
        return res.render('error', { error: err });
    });

    out.on('finish', async function (err) {
        if (err) {
            return res.render('error', { error: err })
        }
        setTimeout(function () {
            mail(req.body.patient_email, "Acupuncture")
        }, 5000);

        res.download(__dirname + '/tmp/receipt.docx', req.body.patient_name + "-TCM.docx")
    })
    docx.generate(out)
});

app.post('/rmt-receipt', function (req, res) {
    var docx = officeGen({
        'type': 'docx',
        'orientation': 'portrait',
        'subject': 'Neshama Therapy Receipt',
        'description': 'Neshama Therapy Receipt'
    });
    docx.on('error', function (err) {
        res.redirect('/');
    });
    let address = '1234 address rd'; //removed for privacy
    if (req.body.new_address) {
        address = req.body.new_address;
    };
    var pObj = docx.createP();
    pObj.options.align = 'left';
    pObj.addImage(path.resolve(__dirname, 'neshamalogo.png'));
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addText(address, { bold: true });
    pObj.addLineBreak();
    pObj.addText('Chen Weinberg, R. Ac, RMT', { bold: true });
    pObj.addLineBreak();
    pObj.addText('RMT Registration number: M658');
    pObj.addLineBreak();
    pObj.addText('College of Massage Therapists of Ontario');
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addHorizontalLine();
    pObj.addText('RMT RECEIPT', { color: '07421e', bold: true, font_size: 14 });
    pObj.addLineBreak();
    pObj.addText('Patient Name: ' + req.body.patient_name, { color: '06a2db', bold: true, font_size: 18 });
    pObj.addLineBreak();
    pObj.addText('Date: ' + req.body.date, { color: '204903', bold: true });
    pObj.addLineBreak();
    pObj.addText('Duration of Treatment: ' + req.body.service_length, { color: '204903', bold: true });
    pObj.addLineBreak();
    pObj.addText('Description Of Service: ' + req.body.service_desc, { color: '204903', bold: true });
    pObj.addLineBreak();
    pObj.addText('Total Amount: ' + req.body.amount, { color: '204903', bold: true });
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addLineBreak();
    pObj.addText('Issued By');
    pObj.addLineBreak();
    pObj.addText('Chen Weinberg, R. Ac, RMT', { color: '10350f', bold: true });
    pObj.addLineBreak();
    pObj.addImage(path.resolve(__dirname, 'signature.JPG'));

    var out = fs.createWriteStream(__dirname + '/tmp/receipt.docx');
    out.on('error', function (err) {
        res.send('Error creating writestream out: \n' + err);
    });
    
    out.on('finish', function (err) {
        if (err) {
            return res.render('error', { error: err })
        }
        setTimeout(function () {
            mail(req.body.patient_email, "RMT")
            .catch((err) => res.send(err));
        }, 5000);

        res.download(__dirname + '/tmp/receipt.docx', req.body.patient_name + "-RMT.docx")
    })
    docx.generate(out)
});


app.get('/download', function (req, res) {
    if (req.downloadName) {
        return res.download('./tmp/receipt.docx', req.downloadName + ".docx");
    }
    res.download('./tmp/receipt.docx');
});

app.listen(process.env.PORT || 3000, function () {
    if (process.env.PORT) {
        console.log("OfficeGenNeshama Running on " + process.env.PORT)
    } else {
        console.log('OfficeGenNeshama running on port 3000!!!')
    }
});