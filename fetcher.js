const Imap = require('imap')
const fs = require('fs')
const simpleParser = require('mailparser').simpleParser
require('dotenv').config()

class MailAttachmentFetcher {
  constructor(emailConfig, localFolderPath) {
    this.emailConfig = emailConfig
    this.localFolderPath = localFolderPath

    if (!fs.existsSync(localFolderPath)) {
      fs.mkdirSync(localFolderPath)
    }

    this.imap = new Imap(emailConfig)
    this.imap.once('ready', this.onImapReady.bind(this))
    this.imap.once('error', this.onImapError.bind(this))
    this.imap.once('end', this.onImapEnd.bind(this))
  }

  onImapReady() {
    console.log('Connection established.')
    this.imap.openBox('INBOX', true, this.onOpenBox.bind(this))
  }

  onImapError(err) {
    console.error('IMAP error:', err)
    throw err
  }

  onImapEnd() {
    console.log('Connection ended.')
  }

  onOpenBox(err) {
    if (err) {
      console.error('Error opening INBOX:', err)
      throw err
    }

    console.log('INBOX opened successfully.')

    this.imap.search(['ALL'], this.onSearchResults.bind(this))
  }

  onSearchResults(searchErr, results) {
    if (searchErr) {
      console.error('Error searching for emails:', searchErr)
      throw searchErr
    }

    console.log(`Fetched ${results.length} emails.`)

    results.forEach((seqno) => {
      const fetch = this.imap.fetch([seqno], { bodies: '' })

      fetch.on('message', (msg, seqno) => {
        msg.on('body', (stream) => {
          simpleParser(stream, (err, parsed) => {
            if (err) {
              console.error('Error parsing email:', err)
              throw err
            }
            if (parsed.attachments && parsed.attachments.length > 0) {
              console.log(`\nSaving attachments from Email ${seqno}`)
              parsed.attachments.forEach((attachment, index) => {
                if (this.isSupportedAttachment(attachment)) {
                  const filename = attachment.filename || `attachment_${seqno}_${index + 1}.${attachment.contentType.split('/')[1]}`
                  const filePath = `${this.localFolderPath}/${filename}`
                  fs.writeFileSync(filePath, attachment.content)

                  console.log(`Saved supported attachment: ${filePath}`)
                } else {
                  console.log(`Skipped unsupported attachment: ${attachment.filename}`)
                }
              })
            } else {
              console.log(`No attachments found in Email ${seqno}.`)
            }
          })
        })
      })
    })

    this.imap.end()
  }

  isSupportedAttachment(attachment) {
    const supportedExtensions = ['.pdf', '.xlsx', '.jpg', '.zip', '.rar', '.docx']

    if (attachment.filename) {
      const extension = attachment.filename.toLowerCase().split('.').pop()
      return supportedExtensions.includes(`.${extension}`)
    }

    return false
  }

  start() {
    this.imap.connect()
  }
}

// Example usage
const emailConfig = {
  user: process.env.EMAIL,
  password: process.env.PASS,
  host: process.env.HOST,
  port: 993,
  tls: true
}

const localFolderPath = './attachments'

const gmailFetcher = new MailAttachmentFetcher(emailConfig, localFolderPath)
gmailFetcher.start()
