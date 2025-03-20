import cors from 'cors'
import dotenv from 'dotenv'
import express from 'express'
import fs from 'fs'
import { google } from 'googleapis'
import http from 'http'
import open from 'open'
import path from 'path'
import puppeteer from 'puppeteer'
import { Server } from 'socket.io'
import { fileURLToPath } from 'url'
import xlsx from 'xlsx'
import logger from './logger.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

dotenv.config()

const app = express()
app.use(
  cors({
    origin: 'http://localhost:3001', // Frontend URL
    methods: ['GET', 'POST'],
    allowedHeaders: ['Content-Type', 'Authorization'],
  })
)

const server = http.createServer(app)
const io = new Server(server, {
  cors: {
    origin: '*',
    methods: ['GET', 'POST'],
  },
})

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, 'public')))

const port = 3000

const CLIENT_ID = process.env.CLIENT_ID
const CLIENT_SECRET = process.env.CLIENT_SECRET
const REDIRECT_URI = `http://localhost:${port}/oauth2callback`
const TOKEN_PATH = 'token.json'

const oauth2Client = new google.auth.OAuth2(
  CLIENT_ID,
  CLIENT_SECRET,
  REDIRECT_URI
)
const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

let verificationCode = ''

// Add this helper function at the top level
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

// Load token if it exists
function loadToken() {
  if (fs.existsSync(TOKEN_PATH)) {
    const token = JSON.parse(fs.readFileSync(TOKEN_PATH))
    oauth2Client.setCredentials(token)
    return true
  }
  return false
}

// Fetch verification code from Gmail
async function getVerificationCode() {
  try {
    const gmail = google.gmail({ version: 'v1', auth: oauth2Client })
    const res = await gmail.users.messages.list({
      userId: 'me',
      maxResults: 5,
    })

    if (!res.data.messages) {
      logger.info('No new emails found.')
      return null
    }

    for (const msg of res.data.messages) {
      const email = await gmail.users.messages.get({
        userId: 'me',
        id: msg.id,
      })
      const body = email.data.snippet
      logger.info('Email body:', body)
      const codeMatch = body.match(/\b\d{6,10}\b/)
      logger.info('Code match:', codeMatch)

      if (codeMatch) {
        return codeMatch[0]
      }
    }

    logger.info('No verification code found in recent emails.')
    return null
  } catch (error) {
    logger.error('Error fetching emails:', error.message)
    return null
  }
}

// Utility function for delays
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
const randomDelay = () => Math.floor(Math.random() * (3000 - 1000) + 1000)

// Add these helper functions for date calculations

// Calculate year and month differences for navigation
const calculateMonthsToNavigate = (
  currentMonth,
  currentYear,
  targetMonth,
  targetYear,
  months
) => {
  // Convert years to months since epoch
  const currentMonthsSinceEpoch =
    parseInt(currentYear) * 12 + months[currentMonth]
  const targetMonthsSinceEpoch = targetYear * 12 + months[targetMonth]

  // The difference is how many months we need to navigate
  return currentMonthsSinceEpoch - targetMonthsSinceEpoch
}

// Modify the date selection part in loginToExpediaPartner function
async function setDateRange(page, start_date, end_date) {
  try {
    // Check if the input dates are already in DD/MM/YYYY format (from formatDateForProcessing)
    const isExpediaFormat = (dateStr) => {
      // If the date is already in DD/MM/YYYY format, the first part should be a day (1-31)
      const parts = dateStr.split('/')
      if (parts.length !== 3) return false
      const firstPart = parseInt(parts[0])
      return firstPart >= 1 && firstPart <= 31 && parts[0].length <= 2
    }

    // Convert from DD/MM/YYYY (Expedia format) to MM/DD/YYYY (internal format) if needed
    const convertToInternalFormat = (dateStr) => {
      if (!isExpediaFormat(dateStr)) return dateStr // Already in MM/DD/YYYY format
      // Input is in DD/MM/YYYY format (Expedia format)
      const [day, month, year] = dateStr.split('/')
      // Return in MM/DD/YYYY format for internal processing
      return `${month}/${day}/${year}`
    }

    // Ensure we're working with internal format (MM/DD/YYYY) for processing
    const internalStartDate = convertToInternalFormat(start_date)
    const internalEndDate = convertToInternalFormat(end_date)

    // Format date with leading zeros - keeping MM/DD/YYYY format for internal use
    const formatDateWithZeros = (dateStr) => {
      // Input is in MM/DD/YYYY format (internal format)
      const [month, day, year] = dateStr.split('/')
      const paddedDay = day.padStart(2, '0')
      const paddedMonth = month.padStart(2, '0')
      console.log('Month:', month)
      console.log('Day:', day)
      console.log('Year:', year)
      console.log('padded month', paddedMonth)
      console.log('padded day', paddedDay)

      // Return in MM/DD/YYYY format to maintain consistency
      return `${paddedMonth}/${paddedDay}/${year}`
    }

    // Format date without leading zeros - keeping MM/DD/YYYY format
    const formatDateWithoutZeros = (dateStr) => {
      // Input is in MM/DD/YYYY format (internal format)
      const [month, day, year] = dateStr.split('/')
      return `${parseInt(month)}/${parseInt(day)}/${year}`
    }

    // Convert input dates to both formats
    const expectedFromDateWithZeros = formatDateWithZeros(internalStartDate)
    const expectedToDateWithZeros = formatDateWithZeros(internalEndDate)
    const expectedFromDateWithoutZeros = formatDateWithoutZeros(internalStartDate)
    const expectedToDateWithoutZeros = formatDateWithoutZeros(internalEndDate)

    logger.info(
      'Converting from date:',
      start_date,
      'to:',
      expectedFromDateWithZeros
    )
    logger.info('Converting to date:', end_date, 'to:', expectedToDateWithZeros)

    // Click the From date input to open calendar
    const fromDateInput = await page.$(
      '.from-input-label input.fds-field-input'
    )
    await fromDateInput.click()
    await new Promise((r) => setTimeout(r, 1000))

    // Step 1: Get current year and month from first calendar
    const firstMonthHeader = await page.$eval('.first-month h2', (el) =>
      el.textContent.trim()
    )
    logger.info('First month header:', firstMonthHeader)
    const [currentMonth, currentYear] = firstMonthHeader.split(' ')
    logger.info('Current month:', currentMonth)
    logger.info('Current year:', currentYear)
    // Convert MM/DD/YYYY to a format JavaScript can parse correctly
    const [month, day, year] = internalStartDate.split('/')
    const targetDate = new Date(`${year}-${month}-${day}`)
    const targetYear = targetDate.getFullYear()
    const targetMonth = targetDate.toLocaleString('en-US', { month: 'long' })
    logger.info('Target month:', targetMonth)
    logger.info('Target year:', targetYear)
    logger.info('Target Date:', targetDate)
    // Validate year
    if (targetYear > parseInt(currentYear)) {
      throw new Error('Target year is greater than current year')
    }

    // Calculate months to navigate for start date
    const totalMonthsToNavigate = calculateMonthsToNavigate(
      currentMonth,
      currentYear,
      targetMonth,
      targetYear,
      {
        January: 1,
        February: 2,
        March: 3,
        April: 4,
        May: 5,
        June: 6,
        July: 7,
        August: 8,
        September: 9,
        October: 10,
        November: 11,
        December: 12,
      }
    )

    logger.info('Total months to navigate (start):', totalMonthsToNavigate)

    // Navigate months for start date
    const navigationButtons = await page.$$('.fds-datepicker-navigation button')
    const navigationButton =
      totalMonthsToNavigate > 0
        ? navigationButtons[0] // prev button for going back in time
        : navigationButtons[1] // next button for going forward

    for (let i = 0; i < Math.abs(totalMonthsToNavigate); i++) {
      await navigationButton.click()
      await new Promise((r) => setTimeout(r, 200))
    }

    // Select day (day - 1 for index)
    const targetDay = targetDate.getDate()
    await page.evaluate((day) => {
      const dayButtons = document.querySelectorAll(
        '.first-month .fds-datepicker-day'
      )
      const dayIndex = day - 1
      if (dayButtons[dayIndex] && !dayButtons[dayIndex].disabled) {
        dayButtons[dayIndex].click()
      }
    }, targetDay)

    // Handle end date selection
    await new Promise((r) => setTimeout(r, 1000))

    // Wait for and click the To date input
    await page.waitForSelector('.to-input-label input.fds-field-input', {
      visible: true,
    })
    const toDateInput = await page.$('.to-input-label input.fds-field-input')
    await page.evaluate((el) => el.click(), toDateInput)
    await new Promise((r) => setTimeout(r, 1000))

    // Make sure calendar is visible before proceeding
    await page.waitForSelector('.second-month h2', { visible: true })

    const secondMonthHeader = await page.$eval('.second-month h2', (el) =>
      el.textContent.trim()
    )
    logger.info('Second month header:', secondMonthHeader)

    const [endCurrentMonth, endCurrentYear] = secondMonthHeader.split(' ')
    logger.info('End current month:', endCurrentMonth)
    logger.info('End current year:', endCurrentYear)
    // Convert MM/DD/YYYY to a format JavaScript can parse correctly
    const [endMonth, endDay, endYear] = internalEndDate.split('/')
    const endDate = new Date(`${endYear}-${endMonth}-${endDay}`)
    const endTargetYear = endDate.getFullYear()
    const endTargetMonth = endDate.toLocaleString('en-US', { month: 'long' })
    logger.info('End target month:', endTargetMonth)
    logger.info('End target year:', endTargetYear)
    logger.info('End Date:', endDate)

    // Calculate months to navigate for end date
    const endTotalMonthsToNavigate = calculateMonthsToNavigate(
      endCurrentMonth,
      endCurrentYear,
      endTargetMonth,
      endTargetYear,
      {
        January: 1,
        February: 2,
        March: 3,
        April: 4,
        May: 5,
        June: 6,
        July: 7,
        August: 8,
        September: 9,
        October: 10,
        November: 11,
        December: 12,
      }
    )

    logger.info('Total end months to navigate:', endTotalMonthsToNavigate)

    // Make sure navigation buttons are visible
    await page.waitForSelector('.fds-datepicker-navigation button', {
      visible: true,
    })
    const navigationButtonsEnd = await page.$$(
      '.fds-datepicker-navigation button'
    )

    // Navigate to end date month
    const endNavigationButton =
      endTotalMonthsToNavigate > 0
        ? navigationButtonsEnd[0] // prev button for going back in time
        : navigationButtonsEnd[1] // next button for going forward

    // Click navigation button with evaluation
    for (let i = 0; i < Math.abs(endTotalMonthsToNavigate); i++) {
      await page.evaluate((button) => button.click(), endNavigationButton)
      await new Promise((r) => setTimeout(r, 200))
    }

    // Select end day with evaluation
    const endTargetDay = endDate.getDate()
    await page.evaluate((day) => {
      const dayButtons = document.querySelectorAll(
        '.second-month .fds-datepicker-day'
      )
      const dayIndex = day - 1
      if (dayButtons[dayIndex] && !dayButtons[dayIndex].disabled) {
        dayButtons[dayIndex].click()
      } else {
        throw new Error('End date day button not found or disabled')
      }
    }, endTargetDay)

    await new Promise((r) => setTimeout(r, 1000))

    // Click done button
    await page.evaluate(() => {
      const doneButton = document.querySelector('.fds-dropdown-footer button')
      if (doneButton) {
        doneButton.click()
      } else {
        throw new Error('Done button not found')
      }
    })
    await new Promise((r) => setTimeout(r, 1000))

    // Verify dates were set
    const fromValue = await page.$eval(
      '.from-input-label input.fds-field-input',
      (el) => el.value
    )
    const toValue = await page.$eval(
      '.to-input-label input.fds-field-input',
      (el) => el.value
    )

    // Convert expected dates to DD/MM/YYYY format for comparison with Expedia's interface
    const convertToExpediaFormat = (dateStr) => {
      // Check if already in DD/MM/YYYY format
      if (isExpediaFormat(dateStr)) return dateStr

      // Input is in MM/DD/YYYY format
      const [month, day, year] = dateStr.split('/')
      // Return in DD/MM/YYYY format as expected by Expedia's interface
      return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`
    }

    const expectedFromDateExpediaFormat = convertToExpediaFormat(start_date)
    const expectedToDateExpediaFormat = convertToExpediaFormat(end_date)

    logger.info(
      'Expected from date (Expedia format):',
      expectedFromDateExpediaFormat
    )
    logger.info('Actual from date:', fromValue)
    logger.info(
      'Expected to date (Expedia format):',
      expectedToDateExpediaFormat
    )
    logger.info('Actual to date:', toValue)

    // Compare with expected dates in Expedia format (DD/MM/YYYY)
    const fromMatches = fromValue === expectedFromDateExpediaFormat
    const toMatches = toValue === expectedToDateExpediaFormat

    if (!fromMatches || !toMatches) {
      logger.info('Date mismatch detected, attempting final correction...')
      await page.evaluate(
        (dates) => {
          const [startDateStr, endDateStr] = dates
          // Parse dates - handle both MM/DD/YYYY and DD/MM/YYYY formats
          let startDay, startMonth, startYear, endDay, endMonth, endYear

          // Check if dates are in DD/MM/YYYY format
          const isExpediaFormat = (dateStr) => {
            const parts = dateStr.split('/')
            if (parts.length !== 3) return false
            const firstPart = parseInt(parts[0])
            return firstPart >= 1 && firstPart <= 31 && parts[0].length <= 2
          }

          if (isExpediaFormat(startDateStr)) {
            // DD/MM/YYYY format
            [startDay, startMonth, startYear] = startDateStr.split('/')
          } else {
            // MM/DD/YYYY format
            [startMonth, startDay, startYear] = startDateStr.split('/')
          }

          if (isExpediaFormat(endDateStr)) {
            // DD/MM/YYYY format
            [endDay, endMonth, endYear] = endDateStr.split('/')
          } else {
            // MM/DD/YYYY format
            [endMonth, endDay, endYear] = endDateStr.split('/')
          }

          // Re-open the date picker
          document
            .querySelector('.from-input-label input.fds-field-input')
            .click()

          // Wait a bit and try to select dates again
          setTimeout(() => {
            // Get all days from both months
            const allDays = document.querySelectorAll('.fds-datepicker-day')

            // Convert to array for easier manipulation
            const daysArray = Array.from(allDays)

            // Find start and end dates considering both months
            const startDateElement = daysArray.find(
              (el) =>
                el.textContent.trim() === startDay &&
                !el.classList.contains('disabled')
            )

            const endDateElement = daysArray.find(
              (el) =>
                el.textContent.trim() === endDay &&
                !el.classList.contains('disabled') &&
                (!startDateElement ||
                  el.compareDocumentPosition(startDateElement) === 4)
            )

            if (startDateElement && endDateElement) {
              startDateElement.click()
              // Add delay between clicks to ensure proper state updates
              setTimeout(() => {
                endDateElement.click()
                // Verify date selection before closing
                setTimeout(() => {
                  const selectedStart = document.querySelector(
                    '.from-input-label input.fds-field-input'
                  ).value
                  const selectedEnd = document.querySelector(
                    '.to-input-label input.fds-field-input'
                  ).value

                  if (selectedStart && selectedEnd) {
                    const doneButton = document.querySelector(
                      '.fds-dropdown-footer button'
                    )
                    if (doneButton) doneButton.click()
                  }
                }, 1000)
              }, 1000)
            }
          }, 1000)
        },
        [start_date, end_date]
      )

      // Wait for the final correction to complete
      await new Promise((r) => setTimeout(r, 2000))
    }

    return { from: fromValue, to: toValue }
  } catch (error) {
    logger.error('Error setting date range:', error)
    throw error
  }
}

// Utility function to split date range into 3-day chunks
function splitDateRange(startDate, endDate) {
  const chunks = []
  // Convert MM/DD/YYYY to a format JavaScript can parse correctly
  const [startMonth, startDay, startYear] = startDate.split('/')
  const [endMonth, endDay, endYear] = endDate.split('/')
  // Use YYYY-MM-DD format for reliable date parsing
  const start = new Date(
    `${startYear}-${startMonth.padStart(2, '0')}-${startDay.padStart(2, '0')}`
  )
  const end = new Date(
    `${endYear}-${endMonth.padStart(2, '0')}-${endDay.padStart(2, '0')}`
  )

  let currentStart = new Date(start)
  while (currentStart < end) {
    let currentEnd = new Date(currentStart)
    currentEnd.setDate(currentEnd.getDate() + 2) // Add 2 days to make it a 3-day chunk

    if (currentEnd > end) {
      currentEnd = new Date(end)
    }

    chunks.push({
      start: currentStart.toLocaleDateString('en-US'),
      end: currentEnd.toLocaleDateString('en-US'),
    })

    currentStart = new Date(currentEnd)
    currentStart.setDate(currentStart.getDate() + 1) // Start next chunk from next day
  }

  return chunks
}

function splitDateRangeIntoChunks(start_date, end_date, chunkSize = 2) {
  // Parse dates correctly regardless of MM/DD/YYYY or DD/MM/YYYY format
  const parseDate = (dateStr) => {
    // Check if the date is in MM/DD/YYYY format (our internal format)
    if (dateStr.includes('/')) {
      const parts = dateStr.split('/')
      // Validate that we have three parts (month, day, year)
      if (
        parts[0].length <= 2 &&
        parts[1].length <= 2 &&
        parts[2].length === 4
      ) {
        // We consistently use MM/DD/YYYY format internally
        // The first part is month, second part is day
        const month = parseInt(parts[0], 10)
        const day = parseInt(parts[1], 10)
        const year = parseInt(parts[2], 10)

        // Create date using reliable YYYY-MM-DD format
        return new Date(
          `${year}-${month.toString().padStart(2, '0')}-${day
            .toString()
            .padStart(2, '0')}`
        )
      }
    }

    // Fallback to default parsing
    return new Date(dateStr)
  }

  // Parse the start and end dates
  const startDate = parseDate(start_date)
  const endDate = parseDate(end_date)

  // Log the parsed dates for debugging
  console.log(`Parsed start date: ${startDate.toISOString()}`)
  console.log(`Parsed end date: ${endDate.toISOString()}`)

  const dateChunks = []
  let currentDate = new Date(startDate)

  // Only create chunks for the specific date range requested
  while (currentDate <= endDate) {
    const nextDate = new Date(currentDate)
    nextDate.setDate(currentDate.getDate() + chunkSize - 1)
    if (nextDate > endDate) {
      nextDate.setDate(endDate.getDate())
    }

    // Format dates as MM/DD/YYYY for internal consistency
    const formatDate = (date) => {
      const month = date.getMonth() + 1
      const day = date.getDate()
      const year = date.getFullYear()
      return `${month}/${day}/${year}`
    }

    dateChunks.push({
      start: formatDate(currentDate),
      end: formatDate(nextDate),
    })

    // Move to next chunk
    currentDate = new Date(nextDate)
    currentDate.setDate(currentDate.getDate() + 1)
  }

  return dateChunks
}

// Puppeteer Login Function
async function loginToExpediaPartner(
  email = process.env.EMAIL,
  password = process.env.PASSWORD,
  start_date = null,
  end_date = null,
  propertyName = null
) {
  let browser = null
  try {
    browser = await puppeteer.launch({
      headless: false,
      defaultViewport: null,
      args: [
        '--start-maximized',
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-web-security',
        '--disable-features=IsolateOrigins,site-per-process',
      ],
      timeout: 60000,
    })

    const page = await browser.newPage()
    await page.setDefaultNavigationTimeout(60000)
    await page.setDefaultTimeout(60000)

    // Navigate to partner central
    logger.info('Navigating to Expedia Partner Central...')
    await page.goto(
      'https://www.expediapartnercentral.com/Account/Logon?signedOff=true',
      {
        waitUntil: ['networkidle0', 'domcontentloaded'],
        timeout: 60000,
      }
    )

    logger.info('Waiting for page load...')

    await delay(randomDelay())

    await page.evaluate(() => {
      window.scrollBy(0, 200) // Scroll down by 200 pixels
    })
    // Wait for email input
    await page.waitForSelector('#emailControl')


    // Type email slowly, character by character
    for (let char of email) {
      await page.type('#emailControl', char, { delay: 100 }) // 100ms delay between each character
    }

    // Click continue button
    await page.click('#continueButton')

    // Wait before entering password
    logger.info('Waiting for password page to load...')

    // Wait for password page to be fully loaded
    try {
      logger.info('Waiting for password page to fully load...')

      // Try to find the password input field with a try-catch to handle both possible selectors
      let passwordInputFound = false

      try {
        // First try to find #password-input
        const passwordInput = await page.waitForSelector('#password-input', {
          visible: true,
          timeout: 15000, // Shorter timeout for first attempt
        })

        if (passwordInput) {
          passwordInputFound = true

          // Add a significant delay to ensure the page is fully loaded and stable
          await delay(3000)

          // Verify the password field is actually ready for input
          const isInputReady = await page.evaluate(() => {
            const input = document.querySelector('#password-input')
            return input && !input.disabled && document.activeElement !== input
          })

          if (!isInputReady) {
            logger.info('Password input not fully ready, waiting longer...')
            await delay(2000)
          }

          // Click on the password field first to ensure focus
          await page.click('#password-input')
          await delay(1000)

          // Clear the field in case there's any text
          await page.evaluate(() => {
            document.querySelector('#password-input').value = ''
          })
          await delay(500)

          logger.info('Password page fully loaded, entering password...')

          // Type password slowly with increased delays
          for (let char of password) {
            await page.type('#password-input', char, { delay: 150 }) // Increased delay
            await delay(100) // Increased delay between characters
          }

          // Wait longer before clicking submit to ensure password is fully entered
          logger.info('Password entered, waiting before clicking submit...')
          await delay(5000)

          // Verify password was entered correctly
          const enteredPassword = await page.evaluate(() => {
            return document.querySelector('#password-input').value
          })

          if (enteredPassword.length !== password.length) {
            logger.warn(
              `Password entry issue: expected ${password.length} chars but got ${enteredPassword.length}`
            )

            // Re-enter password if needed
            await page.evaluate(() => {
              document.querySelector('#password-input').value = ''
            })
            await delay(1000)

            // Try again with even slower typing
            for (let char of password) {
              await page.type('#password-input', char, { delay: 200 })
              await delay(150)
            }
            await delay(2000)
          }

          // Click the login button
          logger.info('Clicking password continue button...')
          await page.click('#password-continue')
        }
      } catch (error) {
        logger.info(
          'Could not find #password-input, trying #passwordControl instead:',
          error.message
        )
        passwordInputFound = false
      }

      // If #password-input wasn't found, try #passwordControl
      if (!passwordInputFound) {
        try {
          // Check if #passwordControl exists
          const passwordControlExists = await page.evaluate(() => {
            return !!document.querySelector('#passwordControl')
          })

          if (!passwordControlExists) {
            logger.info(
              'Neither #password-input nor #passwordControl found. Checking page content...'
            )
            const pageContent = await page.content()
            logger.info('Page title: ' + (await page.title()))
            throw new Error('Password input field not found on the page')
          }

          // Add a significant delay to ensure the page is fully loaded and stable
          await delay(3000)

          // Verify the password field is actually ready for input
          const isInputReady = await page.evaluate(() => {
            const input = document.querySelector('#passwordControl')
            return input && !input.disabled && document.activeElement !== input
          })

          if (!isInputReady) {
            logger.info('Password input not fully ready, waiting longer...')
            await delay(2000)
          }

          // Click on the password field first to ensure focus
          await page.click('#passwordControl')
          await delay(1000)

          // Clear the field in case there's any text
          await page.evaluate(() => {
            document.querySelector('#passwordControl').value = ''
          })
          await delay(500)

          logger.info('Password page fully loaded, entering password...')

          // Type password slowly with increased delays
          for (let char of password) {
            await page.type('#passwordControl', char, { delay: 150 }) // Increased delay
            await delay(100) // Increased delay between characters
          }

          // Wait longer before clicking submit to ensure password is fully entered
          logger.info('Password entered, waiting before clicking submit...')
          await delay(5000)

          // Verify password was entered correctly
          const enteredPassword = await page.evaluate(() => {
            return document.querySelector('#passwordControl').value
          })

          if (enteredPassword.length !== password.length) {
            logger.warn(
              `Password entry issue: expected ${password.length} chars but got ${enteredPassword.length}`
            )

            // Re-enter password if needed
            await page.evaluate(() => {
              document.querySelector('#passwordControl').value = ''
            })
            await delay(1000)

            // Try again with even slower typing
            for (let char of password) {
              await page.type('#passwordControl', char, { delay: 200 })
              await delay(150)
            }
            await delay(2000)
          }

          // Click the login button
          logger.info('Clicking password continue button...')
          await page.click('#signInButton')
        } catch (error) {
          logger.error('Error handling password input:', error.message)
          throw error
        }
      }
    } catch (error) {
      logger.info('Error during password entry:', error.message)
      throw error
    }

    // Wait for verification code page using the correct selector
    logger.info('Waiting for verification page...')
    await page.waitForSelector('input[name="passcode-input"]', {
      visible: true,
      timeout: 60000,
    })

    // Add delay before fetching verification code
    logger.info('Waiting for verification email...')
    await delay(15000) // Wait 15 seconds for email to arrive

    // Get verification code
    const code = await getVerificationCode()
    if (!code) {
      throw new Error('Failed to get verification code from email')
    }
    logger.info('Got verification code:', code)

    // Enter verification code using the correct selector
    await page.type('input[name="passcode-input"]', code, { delay: 100 })
    await delay(randomDelay())

    // Find and click the verify button
    // Try multiple possible selectors since the verify button might have different attributes
    // const verifyButton = await page.evaluate(() => {
    //     const buttons = Array.from(document.querySelectorAll('button'));
    //     const verifyButton = buttons.find(button =>
    //         button.textContent.includes('VERIFY DEVICE') ||
    //         button.textContent.includes('Verify') ||
    //         button.textContent.includes('Submit') ||
    //         button.type === 'submit'
    //     );
    //     return verifyButton;
    // });

    // if (!verifyButton) {
    //     throw new Error('Verify button not found')
    // }
    // await verifyButton.click()
    const verifyButtonHandle = await page.$(
      'button[data-testid="passcode-submit-button"]'
    )

    if (!verifyButtonHandle) {
      throw new Error('Verify button not found')
    }

    // Check if the button is disabled
    const isDisabled = await page.evaluate(
      (button) => button.disabled,
      verifyButtonHandle
    )

    if (isDisabled) {
      throw new Error('Verify button is disabled')
    }

    // Click the button
    await verifyButtonHandle.click()
    logger.info('Clicked the verify button successfully!')

    // Wait for successful login
    await page.waitForNavigation({
      waitUntil: 'networkidle0',
      timeout: 60000,
    })

    logger.info('Login successful!')

    //in here we will check is property name is ser or not
    if (propertyName) {
      // Wait for property table to load
      await page.waitForSelector('.fds-data-table-wrapper', {
        visible: true,
        timeout: 30000,
      })

      // Wait for property search input
      await page.waitForSelector(
        '.all-properties__search input.fds-field-input'
      )

      // Get property name from query params
      logger.info(`Searching for property: ${propertyName}`)

      // Type property name in search
      await page.type(
        '.all-properties__search input.fds-field-input',
        propertyName,
        { delay: 100 }
      )

      // Wait for search results
      await delay(2000)

      // Find and click the property link with more specific selector
      try {
        // Wait for search results to update
        await page.waitForSelector('tbody tr', {
          visible: true,
          timeout: 10000,
        })

        // More specific selector for the property link
        const propertySelector = `.property-cell__property-name a[href*="/lodging/home/home"]`

        const propertyLink = await page.waitForSelector(propertySelector, {
          visible: true,
          timeout: 10000,
        })

        if (propertyLink) {
          // Get the text to verify it's the right property
          const linkText = await page.evaluate(
            (el) => el.textContent,
            propertyLink
          )
          logger.info(`Found property: ${linkText}, clicking...`)

          try {
            // Click the link and wait for navigation
            await Promise.all([
              page.waitForNavigation({
                waitUntil: 'networkidle0',
                timeout: 30000,
              }),
              propertyLink.click(),
            ])

            // Wait for the new page to load
            await delay(8000)
          } catch (error) {
            console.error(error.message)
          }
        }
      } catch (error) {
        logger.error(`Error finding/clicking property: ${error.message}`)
        throw error
      }
    }
    // Find and click the Reservations link
    logger.info('Looking for Reservations link...')

    try {
      // Wait for the drawer content to load
      await page.waitForSelector('.uitk-drawer-content', {
        visible: true,
        timeout: 30000,
      })

      // Click using JavaScript with the exact structure
      const clicked = await page.evaluate(() => {
        const reservationsItem = Array.from(
          document.querySelectorAll('.uitk-action-list-item-content')
        ).find((item) => {
          const textDiv = item.querySelector('.uitk-text.overflow-wrap')
          return textDiv && textDiv.textContent.trim() === 'Reservations'
        })

        if (reservationsItem) {
          const link = reservationsItem.querySelector(
            'a.uitk-action-list-item-link'
          )
          if (link) {
            link.click()
            return true
          }
        }
        return false
      })

      if (!clicked) {
        throw new Error('Could not find or click Reservations link')
      }

      // Wait for navigation to complete
      await Promise.all([
        page.waitForNavigation({
          waitUntil: 'networkidle0',
          timeout: 80000,
        }),
        delay(8000),
      ])

      logger.info('Successfully navigated to Reservations page')

      // Wait for date filters to be visible
      logger.info('Waiting for date filters...')
      await page.waitForSelector('input[type="radio"][name="dateTypeFilter"]', {
        visible: true,
        timeout: 80000,
      })
      ///////////////////////////////////
      //new tab opening
      ///////////////////////////////////
      // Get the current URL
      const currentUrl = page.url()
      console.log(`Current tab URL: ${currentUrl}`)

      // Generate date chunks
      const dateChunks = splitDateRangeIntoChunks(start_date, end_date, 2)
      console.log('Date Chunks:', dateChunks)

      // Process each date chunk sequentially in the same tab
      const allReservations = []

      // Helper function to format dates for Expedia's interface
      const formatDateForProcessing = (dateStr) => {
        // Input is in MM/DD/YYYY format (our internal format)
        const [month, day, year] = dateStr.split('/')
        // Parse date using reliable YYYY-MM-DD format
        const date = new Date(`${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`)
        const formattedDay = date.getDate().toString().padStart(2, '0')
        const formattedMonth = (date.getMonth() + 1).toString().padStart(2, '0')
        const formattedYear = date.getFullYear()
        // Return in DD/MM/YYYY format as expected by Expedia's interface
        return `${formattedDay}/${formattedMonth}/${formattedYear}`
      }

      for (const chunk of dateChunks) {
        logger.info(`Processing chunk: ${chunk.start} to ${chunk.end}`)
        
        // Format dates for URL parameters (YYYY-MM-DD format for Expedia's API)
        const formatDateForUrl = (dateStr) => {
          // Convert MM/DD/YYYY to a format JavaScript can parse correctly
          const [month, day, year] = dateStr.split('/')
          const date = new Date(`${year}-${month}-${day}`)
          const formattedYear = date.getFullYear()
          const formattedMonth = (date.getMonth() + 1).toString().padStart(2, '0')
          const formattedDay = date.getDate().toString().padStart(2, '0')
          return `${formattedYear}-${formattedMonth}-${formattedDay}`
        }

        const startParam = formatDateForUrl(chunk.start)
        const endParam = formatDateForUrl(chunk.end)

        // Construct URL for the current chunk
        let chunkUrl
        if (currentUrl.includes('#')) {
          try {
            const baseUrl = currentUrl.split('#')[0]
            const hash = currentUrl.split('#')[1]
            const params = new URLSearchParams(hash)
            params.set('startDate', startParam)
            params.set('endDate', endParam)
            chunkUrl = `${baseUrl}#${params.toString()}`
          } catch (error) {
            logger.warn(`Error parsing URL hash fragment: ${error.message}`)
            chunkUrl = currentUrl
          }
        } else {
          const separator = currentUrl.includes('?') ? '&' : '?'
          chunkUrl = `${currentUrl}${separator}startDate=${startParam}&endDate=${endParam}`
        }

        // Navigate to the URL for current chunk
        logger.info(`Navigating to: ${chunkUrl}`)
        await page.goto(chunkUrl, { waitUntil: 'networkidle2' })
        await delay(3000) // Wait for page to stabilize

        // Process the current chunk
        const formattedStart = formatDateForProcessing(chunk.start)
        const formattedEnd = formatDateForProcessing(chunk.end)
        
        logger.info(`Processing date range: ${formattedStart} to ${formattedEnd}`)
        const chunkReservations = await processReservationsPage(page, formattedStart, formattedEnd)
        allReservations.push(...chunkReservations)
      }

      logger.info(`Found total ${allReservations.length} reservations across all chunks`)

      // Get current date and time for filename
      const now = new Date()
      const timestamp = now.toISOString().replace(/[:.]/g, '-')

      // Save all reservations to Excel with timestamp
      const workbook = xlsx.utils.book_new()
      const wsData = [
        [
          'Guest Name',
          'Reservation ID',
          'Confirmation Code',
          'Check-in Date',
          'Check-out Date',
          'Room Type',
          'Booking Amount',
          'Booked Date',
          'Card Number',
          'Expiry Date',
          'CVV',
          'Has Card Info',
          'Has Payment Info',
          'Total Guest Payment',
          'Expedia Compensation',
          'Total Payout',
          'Amount to charge/refund',
          'Reason of charge',
          'Status',
        ],
        ...allReservations.map((res) => [
          res.guestName,
          res.reservationId,
          res.confirmationCode,
          res.checkInDate,
          res.checkOutDate,
          res.roomType,
          res.bookingAmount,
          res.bookedDate,
          res.cardNumber || 'N/A',
          res.expiryDate || 'N/A',
          res.cvv || 'N/A',
          res.hasCardInfo ? 'Yes' : 'No',
          res.hasPaymentInfo ? 'Yes' : 'No',
          res.totalGuestPayment || 'N/A',
          res.expediaCompensation || 'N/A',
          res.totalPayout || 'N/A',
          res.amountToChargeOrRefund || 'N/A',
          res.reasonOfCharge || 'N/A',
          res.status || 'Active',
        ]),
      ]

      const ws = xlsx.utils.aoa_to_sheet(wsData)
      xlsx.utils.book_append_sheet(workbook, ws, 'Reservations')
      xlsx.writeFile(workbook, `reservations_${timestamp}.xlsx`)
      logger.info(`Saved reservation data to reservations_${timestamp}.xlsx`)

      // Close the browser
      // await browser.close()
      return allReservations
    } catch (error) {
      logger.error('Error finding/clicking Reservations:', error.message)
      throw error
    }
  } catch (error) {
    logger.error(`Error finding/clicking property: ${error.message}`)
    if (browser) await browser.close()
    throw error
  }
}

// New function to process reservations on a single page
async function processReservationsPage(page, start_date, end_date) {
  try {
    // Click the "Checking out" radio button
    logger.info('Selecting "Checking out" filter...')
    await page.evaluate(() => {
      const radioButtons = Array.from(
        document.querySelectorAll('input[type="radio"][name="dateTypeFilter"]')
      )
      const checkingOutButton = radioButtons.find(
        (radio) =>
          radio.parentElement
            .querySelector('.fds-switch-label')
            .textContent.trim() === 'Checking out'
      )
      if (checkingOutButton) {
        checkingOutButton.click()
      }
    })

    // Wait for radio button click to take effect
    await delay(2000)

    // Set the date range
    logger.info(`Processing date range: ${start_date} to ${end_date}`)
    const dateValues = await setDateRange(page, start_date, end_date)
    logger.info('Set dates:', dateValues)

    //////////////////////////////////////////////////////////////
    //more filter button
    //////////////////////////////////////////////////////////////

    //wait for the more filter button
    logger.info('Waiting for the More filters button...')
    await page.waitForSelector(
      'button.fds-button2.utility.fds-dropdown-trigger',
      {
        visible: true,
        timeout: 10000,
      }
    )

    // Click the More filters button
    await page.evaluate(() => {
      const moreFiltersButton = Array.from(
        document.querySelectorAll(
          'button.fds-button2.utility.fds-dropdown-trigger'
        )
      ).find((button) => {
        const label = button.querySelector('.fds-button2-label')
        return label && label.textContent.trim() === 'More filters'
      })

      if (moreFiltersButton) {
        moreFiltersButton.click()
        return true
      }
      throw new Error('More filters button not found')
    })

    logger.info(
      'Clicked More filters button, waiting for dropdown to appear...'
    )

    // Check the "Expedia Collect Payments" and "Expedia Virtual Card" checkboxes
    await page.evaluate(() => {
      // Find all checkbox labels
      const checkboxLabels = Array.from(
        document.querySelectorAll('.fds-switch-checkbox')
      )

      // Find and click the "Expedia Collect Payments" checkbox
      const expediaCollectPaymentsLabel = checkboxLabels.find(
        (label) =>
          label.querySelector('.fds-switch-label') &&
          label.querySelector('.fds-switch-label').textContent.trim() ===
            'Expedia Collect Payments'
      )

      if (expediaCollectPaymentsLabel) {
        const checkbox = expediaCollectPaymentsLabel.querySelector(
          'input.fds-switch-input'
        )
        if (checkbox && !checkbox.checked) {
          checkbox.click()
          console.log('Checked "Expedia Collect Payments" checkbox')
        }
      } else {
        console.log('Could not find "Expedia Collect Payments" checkbox')
      }

      // Find and click the "Expedia Virtual Card" checkbox
      const expediaVirtualCardLabel = checkboxLabels.find(
        (label) =>
          label.querySelector('.fds-switch-label') &&
          label.querySelector('.fds-switch-label').textContent.trim() ===
            'Expedia Virtual Card'
      )

      if (expediaVirtualCardLabel) {
        const checkbox = expediaVirtualCardLabel.querySelector(
          'input.fds-switch-input'
        )
        if (checkbox && !checkbox.checked) {
          checkbox.click()
          console.log('Checked "Expedia Virtual Card" checkbox')
        }
      } else {
        console.log('Could not find "Expedia Virtual Card" checkbox')
      }
    })

    logger.info(
      "Selected 'Expedia Collect Payments' and 'Expedia Virtual Card' checkboxes"
    )
    await delay(1000) // Wait for checkboxes to be checked

    // Click the Apply button in the dropdown
    await page.evaluate(() => {
      const filterApplyButton = Array.from(
        document.querySelectorAll(
          '.fds-dropdown-actions button.fds-button2.utility'
        )
      ).find((button) => {
        const label = button.querySelector('.fds-button2-label')
        return label && label.textContent.trim() === 'Apply'
      })

      if (filterApplyButton) {
        filterApplyButton.click()
        return true
      }
    })

    logger.info('Applied filters from dropdown')
    await delay(2000)

    logger.info('Waiting for data to load...')

    // Wait for the loading indicator to appear
    await page
      .waitForSelector('td .fds-loader.is-loading.is-visible', {
        visible: true,
        timeout: 10000,
      })
      .catch(() => logger.info('Loading indicator did not appear'))

    // Wait for the loading indicator to disappear
    await page.waitForSelector('td .fds-loader.is-loading.is-visible', {
      hidden: true,
      timeout: 30000,
    })

    logger.info('Loading completed, continuing with data processing...')

    // Then continue with your existing code for processing the data...
    logger.info('Starting to process reservation data...')

    // Wait for the table to be visible
    await page.waitForSelector('table.fds-data-table', {
      visible: true,
      timeout: 30000,
    })

    // Wait for data to load and stabilize
    let previousCount = 0
    let attempts = 0
    const maxAttempts = 15 // Increased max attempts

    while (attempts < maxAttempts) {
      await delay(2000)

      const currentCount = await page.evaluate(() => {
        return document.querySelectorAll('td.guestName button.guestNameLink')
          .length
      })

      logger.info(
        `Found ${currentCount} reservations on attempt ${attempts + 1}...`
      )

      if (currentCount === previousCount && currentCount > 0) {
        logger.info('Data count stabilized')
        break
      }

      previousCount = currentCount
      attempts++
    }

    // Final verification
    const finalCount = await page.evaluate(() => {
      return document.querySelectorAll('td.guestName button.guestNameLink')
        .length
    })

    logger.info(`Final reservation count: ${finalCount}`)

    if (finalCount === 0) {
      logger.info('No reservations found after multiple attempts')
      return []
    }

    // After date range is applied and before scraping data
    logger.info('Setting results per page to 100...')
    await page.waitForSelector('.fds-pagination-selector select')
    await page.click('.fds-pagination-selector select')
    await page.select('.fds-pagination-selector select', '100')

    // Wait for data to reload with 100 records
    await delay(3000)
    await page.waitForSelector('table.fds-data-table tbody tr', {
      visible: true,
      timeout: 30000,
    })

    // Initialize array for all reservations with Set for tracking duplicates
    const pageReservations = []
    const processedReservationIds = new Set()

    // Function to check if there's a next page
    const hasNextPage = async () => {
      return await page.evaluate(() => {
        const nextButton = document.querySelector(
          '.fds-pagination-button.next button'
        )
        return nextButton && !nextButton.disabled
      })
    }

    // Function to get total results count
    const getTotalResults = async () => {
      const resultsText = await page.$eval(
        '.fds-pagination-showing-result',
        (el) => el.textContent
      )
      const match = resultsText.match(/of (\d+) Results/)
      return match ? parseInt(match[1]) : 0
    }

    const totalResults = await getTotalResults()
    logger.info(`Total reservations to fetch: ${totalResults}`)

    let currentPage = 1
    let hasMore = true

    while (hasMore) {
      try {
        logger.info(`Processing page ${currentPage}...`)

        // Wait for table data to load
        await page.waitForSelector('table.fds-data-table tbody tr', {
          visible: true,
          timeout: 30000,
        })
        await delay(5000)

        // Get reservations from current page
        const rows = await page.$$('table.fds-data-table tbody tr')

        for (const row of rows) {
          try {
            // Get basic data first
            const basicData = await page.evaluate((row) => {
              return {
                guestName:
                  row
                    .querySelector(
                      'td.guestName button.guestNameLink span.fds-button2-label'
                    )
                    ?.textContent.trim() || '',
                reservationId:
                  row
                    .querySelector('td.reservationId div.fds-cell')
                    ?.textContent.trim() || '',
                confirmationCode:
                  row
                    .querySelector(
                      'td.confirmationCode label.confirmationCodeLabel'
                    )
                    ?.textContent.trim() || '',
                checkInDate:
                  row.querySelector('td.checkInDate')?.textContent.trim() || '',
                checkOutDate:
                  row.querySelector('td.checkOutDate')?.textContent.trim() ||
                  '',
                roomType:
                  row.querySelector('td.roomType')?.textContent.trim() || '',
                bookingAmount:
                  row
                    .querySelector('td.bookingAmount .fds-currency-value')
                    ?.textContent.trim() || '',
                bookedDate:
                  row.querySelector('td.bookedOnDate')?.textContent.trim() ||
                  '',
              }
            }, row)

            // Check if we've already processed this reservation
            if (processedReservationIds.has(basicData.reservationId)) {
              logger.info(
                `Skipping duplicate reservation: ${basicData.reservationId}`
              )
              continue
            }

            // Add to processed set
            processedReservationIds.add(basicData.reservationId)

            // Get card details
            const guestNameButton = await row.$(
              'td.guestName button.guestNameLink'
            )
            await guestNameButton.click()

            // Wait for initial dialog to appear with timeout
            try {
              await Promise.race([
                page.waitForSelector('.fds-dialog', {
                  visible: true,
                  timeout: 8000,
                }),
                new Promise((_, reject) =>
                  setTimeout(() => reject(new Error('Dialog timeout')), 8000)
                ),
              ])
            } catch (error) {
              logger.info(
                'Dialog did not appear within timeout, skipping to next reservation'
              )
              continue
            }

            // Check if this is a canceled reservation
            const isCanceled = await page.evaluate(() => {
              const dialogTitle = document.querySelector('.fds-dialog-title')
              return (
                dialogTitle && dialogTitle.textContent.includes('Cancelled')
              )
            })

            if (isCanceled) {
              logger.info('Found canceled reservation, closing dialog...')
              try {
                // Try multiple methods to close the dialog
                await Promise.race([
                  // Method 1: Click the close button
                  page.click('.fds-dialog-header button.dialog-close'),
                  // Method 2: Use JavaScript to click the close button
                  page.evaluate(() => {
                    const closeButton = document.querySelector(
                      '.fds-dialog-header button.dialog-close'
                    )
                    if (closeButton) closeButton.click()
                  }),
                  // Method 3: Press Escape key
                  page.keyboard.press('Escape'),
                ])
                await delay(1500) // Wait for dialog to close
                continue // Skip to next reservation
              } catch (error) {
                logger.warn(
                  'Warning: Could not close canceled reservation dialog'
                )
                continue
              }
            }

            // Wait a bit for content to load
            await delay(2000)

            // Scroll to the bottom of dialog content and wait
            await page.evaluate(() => {
              const dialogContent = document.querySelector(
                '.fds-dialog-content'
              )
              if (dialogContent) {
                dialogContent.scrollTo(0, dialogContent.scrollHeight)
              }
            })

            // Wait for content to load after scroll
            await delay(2000)

            // Get card details with retry mechanism
            let cardData = null
            let paymentData = null
            let remainingAmountToCharge = null
            let amountToRefund = null
            let retries = 0
            while (!cardData && !paymentData && retries < 3) {
              try {
                // First try to get card details
                cardData = await page.evaluate(() => {
                  const cardNumber =
                    document
                      .querySelector('.cardNumber.replay-conceal bdi')
                      ?.textContent.trim() || ''
                  const expiryDate =
                    document
                      .querySelector(
                        '.cardDetails .fds-cell.all-cell-1-4.fds-type-color-primary.replay-conceal'
                      )
                      ?.textContent.trim() || ''
                  const cvv =
                    document
                      .querySelectorAll(
                        '.cardDetails .fds-cell.all-cell-1-4.fds-type-color-primary.replay-conceal'
                      )[1]
                      ?.textContent.trim() || ''

                  if (cardNumber) {
                    return {
                      cardNumber,
                      expiryDate,
                      cvv,
                    }
                  }
                  return null
                })

                // If no card data, try to get payment information
                if (!cardData) {
                  paymentData = await page.evaluate(() => {
                    // Find all section titles
                    const sectionTitles = Array.from(
                      document.querySelectorAll('.sidePanelSectionTitle')
                    )

                    // Find the payment sections
                    const totalGuestPaymentTitle = sectionTitles.find((el) =>
                      el.textContent.includes('Total guest payment')
                    )
                    const expediaCompensationTitle = sectionTitles.find((el) =>
                      el.textContent.includes('Expedia compensation')
                    )
                    const totalPayoutTitle = sectionTitles.find((el) =>
                      el.textContent.includes('Your total payout')
                    )

                    // Get the values
                    const totalGuestPayment =
                      totalGuestPaymentTitle?.nextElementSibling
                        ?.querySelector('.fds-currency-value')
                        ?.textContent.trim() || ''
                    const expediaCompensation =
                      expediaCompensationTitle?.nextElementSibling
                        ?.querySelector('.fds-currency-value')
                        ?.textContent.trim() || ''
                    const totalPayout =
                      totalPayoutTitle?.nextElementSibling
                        ?.querySelector('.fds-currency-value')
                        ?.textContent.trim() || ''

                    if (totalGuestPayment) {
                      return {
                        totalGuestPayment,
                        expediaCompensation,
                        totalPayout,
                      }
                    }
                    return null
                  })
                }

                // Extract "Remaining amount to charge" and "Amount to refund"
                const additionalPaymentInfo = await page.evaluate(() => {
                  // Find "Remaining amount to charge"
                  const remainingAmountSection = Array.from(
                    document.querySelectorAll('.fds-cell.sidePanelSection')
                  ).find((section) =>
                    section.textContent.includes('Remaining amount to charge')
                  )

                  const remainingAmount =
                    remainingAmountSection
                      ?.querySelector('.fds-currency-value')
                      ?.textContent.trim() || ''

                  // Find "Amount to refund"
                  const refundSection = Array.from(
                    document.querySelectorAll('.fds-grid.sidePanelSection')
                  ).find((section) =>
                    section.textContent.includes('Amount to refund')
                  )

                  const refundAmount =
                    refundSection
                      ?.querySelector('.fds-currency-value')
                      ?.textContent.trim() || ''

                  return {
                    remainingAmountToCharge: remainingAmount,
                    amountToRefund: refundAmount,
                  }
                })

                if (additionalPaymentInfo) {
                  remainingAmountToCharge =
                    additionalPaymentInfo.remainingAmountToCharge
                  amountToRefund = additionalPaymentInfo.amountToRefund

                  if (remainingAmountToCharge) {
                    logger.info(
                      `Found Remaining amount to charge: ${remainingAmountToCharge}`
                    )
                  }

                  if (amountToRefund) {
                    logger.info(`Found Amount to refund: ${amountToRefund}`)
                  }
                }
              } catch (e) {
                retries++
                await delay(1000)
              }
            }

            //////////////////////////////////////////////////////////////
            //close the side panel
            //////////////////////////////////////////////////////////////
            try {
              await page.click('.fds-dialog-header button.dialog-close')
              await delay(1500)
            } catch (e) {
              logger.warn('Warning: Could not close dialog normally')
            }

            // Add to reservations array with either card data or payment data
            pageReservations.push({
              ...basicData,
              ...(cardData || {}),
              ...(paymentData || {}),
              hasCardInfo: !!cardData,
              hasPaymentInfo: !!paymentData,
              remainingAmountToCharge: remainingAmountToCharge || 'N/A',
              amountToRefund: amountToRefund || 'N/A',
              amountToChargeOrRefund:
                remainingAmountToCharge || amountToRefund || 'N/A',
              reasonOfCharge: remainingAmountToCharge
                ? 'Remaining Amount to Charge'
                : amountToRefund
                ? 'Amount to Refund'
                : 'N/A',
            })
          } catch (error) {
            logger.info(`Error processing reservation: ${error.message}`)
            if (basicData) {
              pageReservations.push({
                ...basicData,
                cardNumber: 'N/A',
                expiryDate: 'N/A',
                cvv: 'N/A',
                remainingAmountToCharge: 'N/A',
                amountToRefund: 'N/A',
                amountToChargeOrRefund: 'N/A',
                reasonOfCharge: 'N/A',
              })
            }
          }
        }

        logger.info(
          `Processed ${pageReservations.length} of ${totalResults} reservations`
        )

        // Check if there's a next page
        hasMore = await hasNextPage()
        if (hasMore) {
          await page.click('.fds-pagination-button.next button')
          await delay(2000)
          currentPage++
        }
      } catch (pageError) {
        logger.info(
          `Error processing page ${currentPage}: ${pageError.message}`
        )
        // Try to recover by reloading the page
        await page.reload({ waitUntil: 'networkidle0' })
        await delay(5000)
      }
    }

    logger.info(
      `Found total ${pageReservations.length} reservations on this tab`
    )
    return pageReservations
  } catch (error) {
    logger.error(`Error processing tab: ${error.message}`)
    return []
  }
}

// API endpoint to get logs
app.get('/api/data', (req, res) => {
  try {
    const data = JSON.parse(
      fs.readFileSync(path.join(__dirname, 'data.json'), 'utf8')
    )
    res.json(data)
  } catch (error) {
    res.status(500).json({ error: 'Error reading logs' })
  }
})

// Watch for JSON file changes
fs.watch(path.join(__dirname, 'data.json'), () => {
  try {
    const data = JSON.parse(
      fs.readFileSync(path.join(__dirname, 'data.json'), 'utf8')
    )
    io.emit('update', data) // Broadcast updates
  } catch (error) {
    console.error('Error reading logs:', error)
  }
})

// WebSocket connection handling
io.on('connection', (socket) => {
  console.log('Client connected')

  fs.readFile(path.join(__dirname, 'data.json'), 'utf8', (err, data) => {
    if (!err) socket.emit('update', JSON.parse(data)) // Send initial data
  })

  socket.on('disconnect', () => console.log('Client disconnected'))
})

// Express routes
app.get('/auth', async (req, res) => {
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  })

  res.redirect(authUrl)
})

app.get('/oauth2callback', async (req, res) => {
  const code = req.query.code
  if (!code) {
    return res.status(400).send('Authorization code not found.')
  }

  try {
    const { tokens } = await oauth2Client.getToken(code)
    oauth2Client.setCredentials(tokens)
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens))
    // res.send('Authentication successful! You can close this window.')
    res.redirect(process.env.FRONTEND_REDIRECT_URI)
  } catch (error) {
    res.status(500).send('Error retrieving access token: ' + error.message)
  }
})
// Independent API endpoint for Expedia login automation
app.get('/api/expedia', async (req, res) => {
  const { email, password, start_date, end_date, propertyName } = req.query

  if (!email || !password || !start_date || !end_date) {
    return res.status(400).json({
      success: false,
      message: 'Email, password, start_date, and end_date are required',
    })
  }

  try {
    if (!loadToken()) {
      return res
        .status(401)
        .json({ success: false, message: 'Gmail authentication required' })
    }

    // Log the original dates for debugging
    console.log('Original start date:', start_date)
    console.log('Original end date:', end_date)

    // Keep dates in MM/DD/YYYY format for consistency
    // This ensures the date parsing is consistent throughout the application
    const validateDateFormat = (dateStr) => {
      const parts = dateStr.split('/')
      // Ensure we have three parts (month, day, year)
      if (parts.length !== 3) {
        throw new Error(`Invalid date format: ${dateStr}. Expected MM/DD/YYYY`)
      }

      // Parse the parts
      const month = parseInt(parts[0], 10)
      const day = parseInt(parts[1], 10)
      const year = parseInt(parts[2], 10)

      // Basic validation
      if (
        isNaN(month) ||
        isNaN(day) ||
        isNaN(year) ||
        month < 1 ||
        month > 12 ||
        day < 1 ||
        day > 31
      ) {
        throw new Error(`Invalid date values in: ${dateStr}`)
      }

      return dateStr
    }

    // Validate the date formats
    const validatedStartDate = validateDateFormat(start_date)
    const validatedEndDate = validateDateFormat(end_date)

    logger.info('Validated start date:', validatedStartDate)
    logger.info('Validated end date:', validatedEndDate)

    // Call loginToExpediaPartner with the validated dates
    await loginToExpediaPartner(
      email,
      password,
      validatedStartDate,
      validatedEndDate,
      propertyName
    )

    res.json({
      success: true,
      message: 'Successfully processed',
    })
  } catch (error) {
    res.status(500).json({ success: false, message: error.message })
  }
})

// app.get('/oauth2callback', async (req, res) => {
//   const code = req.query.code
//   if (!code) {
//     return res.status(400).send('Authorization code not found.')
//   }

//   try {
//     const { tokens } = await oauth2Client.getToken(code)
//     oauth2Client.setCredentials(tokens)
//     fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens))
//     res.send('Authentication successful! You can close this window.')
//     await loginToExpediaPartner('ghbahamar@epchotels.com', 'Ritjiavik2010$')
//   } catch (error) {
//     res.status(500).send('Error retrieving access token: ' + error.message)
//   }
// })

// Start the Express server
server.listen(port, () => {
  logger.info(`Server running at http://localhost:${port}`)
  if (!loadToken()) {
    logger.info('Opening browser for authentication...')
    open(`http://localhost:${port}/auth`)
  }
})
