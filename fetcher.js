const Imap = require('imap')
const fs = require('fs')
const simpleParser = require('mailparser').simpleParser
require('dotenv').config()
const inspect = require('util').inspect

/**
 * MailAttachmentFetcher class
 * @class
 * @description Fetches email attachments from an IMAP server and saves them to a local folder.
 */
class MailAttachmentFetcher {
  /**
   * Constructor
   * @param {Object} emailConfig - Email configuration object.
   * @param {string} localFolderPath - Local folder path to save attachments.
   * @param {string} scinceDate - Date since which to fetch emails (e.g., 'Jan 1, 2024').
   * @example
   * const emailConfig = {
   *   user: 'example@gmail.com',
   *   password: 'password123',
   *   host: 'imap.gmail.com',
   *   port: 993,
   *   tls: true
   * };
   * const fetcher = new MailAttachmentFetcher(emailConfig, './attachments', 'Jan 1, 2024');
   */
  constructor(emailConfig, localFolderPath, scinceDate) {
    this.scinceDate = scinceDate
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

  /**
   * Called when IMAP connection is established.
   * @private
   * @returns {void}
   */
  onImapReady() {
    console.log('Connection established.')
    this.imap.openBox('INBOX', true, this.onOpenBox.bind(this))
  }

  /**
   * Called when an IMAP error occurs.
   * @param {Error} err - Error object.
   * @private
   * @returns {void}
   */
  onImapError(err) {
    console.error('IMAP error:', err)
    throw err
  }

  /**
   * Called when the IMAP connection ends.
   * @private
   * @returns {void}
   */
  onImapEnd() {
    console.log('Connection ended.')
  }

  /**
   * Called when the mailbox is successfully opened.
   * @param {Error|null} err - Error object, or null if successful.
   * @private
   * @returns {void}
   */
  onOpenBox(err) {
    if (err) {
      console.error('Error opening INBOX:', err)
      throw err
    }

    console.log('INBOX opened successfully.')
    this.imap.search(['ALL', ['SINCE', this.scinceDate]], this.onSearchResults.bind(this))
  }

  /**
   * Called when search results are received.
   * @param {Error|null} searchErr - Error object, or null if successful.
   * @param {Array<number>} results - Array of email sequence numbers.
   * @private
   * @returns {void}
   */
  onSearchResults(searchErr, results) {
    if (searchErr) {
      console.error('Error searching for emails:', searchErr)
      throw searchErr
    }

    console.log(`Fetched ${results.length} emails since ${this.scinceDate}.`)
    const existingAttachments = fs.readdirSync(this.localFolderPath)

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
              parsed.attachments.forEach((attachment, index) => {
                if (this.isSupportedAttachment(attachment)) {
                  // console.log(inspect(parsed.attachments[index]))
                  const filename = attachment.filename || `attachment_${seqno}_${index + 1}.${attachment.contentType.split('/')[1]}`
                  if (existingAttachments.includes(filename)) {
                    console.log(`Attachment ${filename} already exists. Skipping...`)
                    return
                  }

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

  /**
   * Checks if an attachment is supported.
   * @param {Object} attachment - Attachment object containing filename and contentType.
   * @returns {boolean} True if the attachment is supported, false otherwise.
   * @example
   * const attachment = { filename: 'document.pdf', contentType: 'application/pdf' };
   * console.log(fetcher.isSupportedAttachment(attachment)); // true
   */
  isSupportedAttachment(attachment) {
    const supportedExtensions = ['.pdf', '.xlsx', '.jpg', '.zip', '.rar', '.docx']
    if (attachment.filename) {
      const extension = attachment.filename.toLowerCase().split('.').pop()
      return supportedExtensions.includes(`.${extension}`)
    }
    return false
  }

  /**
   * Starts the IMAP connection.
   * @public
   * @returns {void}
   * @example
   * fetcher.start();
   */
  start() {
    this.imap.connect()
  }
}

const emailConfig = {
  user: process.env.EMAIL,
  password: process.env.PASS,
  host: process.env.HOST,
  port: 993,
  tls: true,
  tlsOptions: { rejectUnauthorized: false }
}

const localFolderPath = './attachments'

const gmailFetcher = new MailAttachmentFetcher(emailConfig, localFolderPath, 'Nov 24, 2024')
gmailFetcher.start()
