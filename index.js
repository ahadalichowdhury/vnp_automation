Final

import express from 'express'
import fs from 'fs'
import { google } from 'googleapis'
import open from 'open'
import puppeteer from 'puppeteer'
import xlsx from 'xlsx'
import dotenv from 'dotenv'

dotenv.config()

const app = express()
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
const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

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
    const res = await gmail.users.messages.list({ userId: 'me', maxResults: 5 })

    if (!res.data.messages) {
      console.log('No new emails found.')
      return null
    }

    for (const msg of res.data.messages) {
      const email = await gmail.users.messages.get({ userId: 'me', id: msg.id })
      const body = email.data.snippet
      console.log('Email body:', body)
      const codeMatch = body.match(/\b\d{6,10}\b/)
      console.log('Code match:', codeMatch)

      if (codeMatch) {
        return codeMatch[0]
      }
    }

    console.log('No verification code found in recent emails.')
    return null
  } catch (error) {
    console.error('Error fetching emails:', error.message)
    return null
  }
}

// Utility function for delays
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
const randomDelay = () => Math.floor(Math.random() * (3000 - 1000) + 1000)

// Add these helper functions for date calculations
function calculateMonthIndex(currentMonth, targetMonth) {
  // Convert month names to numbers (1-12)
  const months = {
    'January': 1, 'February': 2, 'March': 3, 'April': 4,
    'May': 5, 'June': 6, 'July': 7, 'August': 8,
    'September': 9, 'October': 10, 'November': 11, 'December': 12
  }
  
  let current = months[currentMonth]
  let target = months[targetMonth]
  
  // Calculate difference
  return current - target
}

function calculateDateDifference(currentDate, targetDate) {
  const current = new Date(currentDate)
  const target = new Date(targetDate)
  
  // Calculate year difference * 12 for months
  const yearDiff = (current.getFullYear() - target.getFullYear()) * 12
  
  // Add month difference
  const monthDiff = current.getMonth() - target.getMonth()
  
  return yearDiff + monthDiff
}

// Modify the date selection part in loginToExpediaPartner function
async function setDateRange(page, start_date, end_date) {
  try {
    // Format date with leading zeros
    const formatDateWithZeros = (dateStr) => {
      const [month, day, year] = dateStr.split('/')
      const paddedDay = day.padStart(2, '0')
      const paddedMonth = month.padStart(2, '0')
      return `${paddedDay}/${paddedMonth}/${year}`
    }

    // Format date without leading zeros
    const formatDateWithoutZeros = (dateStr) => {
      const [month, day, year] = dateStr.split('/')
      return `${parseInt(day)}/${parseInt(month)}/${year}`
    }

    // Convert input dates to both formats
    const expectedFromDateWithZeros = formatDateWithZeros(start_date)
    const expectedToDateWithZeros = formatDateWithZeros(end_date)
    const expectedFromDateWithoutZeros = formatDateWithoutZeros(start_date)
    const expectedToDateWithoutZeros = formatDateWithoutZeros(end_date)

    console.log('Converting from date:', start_date, 'to:', expectedFromDateWithZeros)
    console.log('Converting to date:', end_date, 'to:', expectedToDateWithZeros)

    // Click the From date input to open calendar
    const fromDateInput = await page.$('.from-input-label input.fds-field-input')
    await fromDateInput.click()
    await new Promise(r => setTimeout(r, 1000))

    // Step 1: Get current year and month from first calendar
    const firstMonthHeader = await page.$eval('.first-month h2', el => el.textContent.trim())
    console.log('First month header:', firstMonthHeader)
    const [currentMonth, currentYear] = firstMonthHeader.split(' ')
    console.log('Current month:', currentMonth)
    console.log('Current year:', currentYear)
    const targetDate = new Date(start_date)
    const targetYear = targetDate.getFullYear()
    const targetMonth = targetDate.toLocaleString('en-US', { month: 'long' })
    console.log('Target month:', targetMonth)
    console.log('Target year:', targetYear)
    console.log('Target Date:', targetDate)
    // Validate year
    if (targetYear > parseInt(currentYear)) {
      throw new Error('Target year is greater than current year')
    }

    // Calculate year difference (fixed)
    const yearDiff = (targetYear - parseInt(currentYear)) * 12
    console.log('Year difference:', yearDiff)

    // Step 2: Calculate month index difference (fixed)
    const months = {
      'January': 1, 'February': 2, 'March': 3, 'April': 4,
      'May': 5, 'June': 6, 'July': 7, 'August': 8,
      'September': 9, 'October': 10, 'November': 11, 'December': 12
    }
    
    let monthDiff = 0
    if (currentMonth !== targetMonth) {
      monthDiff = months[targetMonth] - months[currentMonth]
    }

    // Total navigation needed
    const totalMonthsToNavigate = yearDiff + monthDiff
    console.log('Total months to navigate (start):', totalMonthsToNavigate)

    // Navigate months
    const navigationButtons = await page.$$('.fds-datepicker-navigation button')
    const navigationButton = totalMonthsToNavigate > 0 
      ? navigationButtons[1] // next button
      : navigationButtons[0] // prev button
    
    for (let i = 0; i < Math.abs(totalMonthsToNavigate); i++) {
      await navigationButton.click()
      await new Promise(r => setTimeout(r, 200))
    }

    // Select day (day - 1 for index)
    const targetDay = targetDate.getDate()
    await page.evaluate((day) => {
      const dayButtons = document.querySelectorAll('.first-month .fds-datepicker-day')
      const dayIndex = day - 1
      if (dayButtons[dayIndex] && !dayButtons[dayIndex].disabled) {
        dayButtons[dayIndex].click()
      }
    }, targetDay)

    // Handle end date selection - Updated with better waiting
    await new Promise(r => setTimeout(r, 1000))
    
    // Wait for and click the To date input
    await page.waitForSelector('.to-input-label input.fds-field-input', { visible: true })
    const toDateInput = await page.$('.to-input-label input.fds-field-input')
    await page.evaluate(el => el.click(), toDateInput)
    await new Promise(r => setTimeout(r, 1000))

    // Make sure calendar is visible before proceeding
    await page.waitForSelector('.second-month h2', { visible: true })
    
    // Get second month header for end date
    const secondMonthHeader = await page.$eval('.second-month h2', el => el.textContent.trim())
    console.log('Second month header:', secondMonthHeader)
    const [endCurrentMonth, endCurrentYear] = secondMonthHeader.split(' ')
    console.log('End current month:', endCurrentMonth)
    console.log('End current year:', endCurrentYear)
    const endDate = new Date(end_date)
    const endTargetYear = endDate.getFullYear()
    const endTargetMonth = endDate.toLocaleString('en-US', { month: 'long' })
    console.log('End target month:', endTargetMonth)
    console.log('End target year:', endTargetYear)
    console.log('End Date:', endDate)
    // Calculate end date navigation (fixed)
    const endYearDiff = (endTargetYear - parseInt(endCurrentYear)) * 12
    console.log('End year difference:', endYearDiff)
    
    let endMonthDiff = 0
    if (endCurrentMonth !== endTargetMonth) {
      endMonthDiff = months[endTargetMonth] - months[endCurrentMonth]
    }
    console.log('End month difference:', endMonthDiff)

    const endTotalMonthsToNavigate = endYearDiff + endMonthDiff
    console.log('Total months to navigate (end):', endTotalMonthsToNavigate)

    // Make sure navigation buttons are visible
    await page.waitForSelector('.fds-datepicker-navigation button', { visible: true })
    const navigationButtonsEnd = await page.$$('.fds-datepicker-navigation button')
    
    // Navigate to end date month
    const endNavigationButton = endTotalMonthsToNavigate > 0
      ? navigationButtonsEnd[1]
      : navigationButtonsEnd[0]

    // Click navigation button with evaluation
    for (let i = 0; i < Math.abs(endTotalMonthsToNavigate); i++) {
      await page.evaluate(button => button.click(), endNavigationButton)
      await new Promise(r => setTimeout(r, 200))
    }

    // Wait for days to be visible before selecting
    await page.waitForSelector('.second-month .fds-datepicker-day', { visible: true })
    
    // Select end day with evaluation
    const endTargetDay = endDate.getDate()
    await page.evaluate((day) => {
      const dayButtons = document.querySelectorAll('.second-month .fds-datepicker-day')
      const dayIndex = day - 1
      if (dayButtons[dayIndex] && !dayButtons[dayIndex].disabled) {
        dayButtons[dayIndex].click()
      } else {
        throw new Error('End date day button not found or disabled')
      }
    }, endTargetDay)

    // Try clicking Done button directly
    await page.evaluate(() => {
      const doneButton = document.querySelector('.fds-dropdown-footer button')
      if (doneButton) {
        doneButton.click()
      } else {
        throw new Error('Done button not found')
      }
    })
    await new Promise(r => setTimeout(r, 1000))

    // Format dates for comparison
    const formatDate = (dateStr) => {
      // Split the date string into parts
      const [month, day, year] = dateStr.split('/')
      return `${month}/${day}/${year}`
    }

    // Parse dates correctly
    const parseDate = (dateStr) => {
      const [month, day, year] = dateStr.split('/')
      return new Date(year, month - 1, day) // month is 0-based in Date constructor
    }

    // Wait a bit for dates to be set
    await new Promise(r => setTimeout(r, 2000))

    // Verify dates were set
    const fromValue = await page.$eval('.from-input-label input.fds-field-input', el => el.value)
    const toValue = await page.$eval('.to-input-label input.fds-field-input', el => el.value)

    console.log('Expected from date (with zeros):', expectedFromDateWithZeros)
    console.log('Expected from date (without zeros):', expectedFromDateWithoutZeros)
    console.log('Actual from date:', fromValue)
    console.log('Expected to date (with zeros):', expectedToDateWithZeros)
    console.log('Expected to date (without zeros):', expectedToDateWithoutZeros)
    console.log('Actual to date:', toValue)

    // Compare with expected dates, accepting either format
    const fromMatches = fromValue === expectedFromDateWithZeros || fromValue === expectedFromDateWithoutZeros
    const toMatches = toValue === expectedToDateWithZeros || toValue === expectedToDateWithoutZeros

    if (!fromMatches || !toMatches) {
      console.log('Date mismatch detected')
      // Try clicking the dates again
      await page.evaluate((dates) => {
        const [startDateStr, endDateStr] = dates
        const [startMonth, startDay, startYear] = startDateStr.split('/')
        const [endMonth, endDay, endYear] = endDateStr.split('/')
        
        // Find and click start date
        const startButtons = document.querySelectorAll('.first-month .fds-datepicker-day')
        for (const button of startButtons) {
          if (button.textContent.trim() === startDay) {
            button.click()
            break
          }
        }
        
        // Find and click end date
        const endButtons = document.querySelectorAll('.second-month .fds-datepicker-day')
        for (const button of endButtons) {
          if (button.textContent.trim() === endDay) {
            button.click()
            break
          }
        }

        // Click done button
        const doneButton = document.querySelector('.fds-dropdown-footer button')
        if (doneButton) doneButton.click()
      }, [start_date, end_date])

      // Wait and verify again
      await new Promise(r => setTimeout(r, 2000))
      const newFromValue = await page.$eval('.from-input-label input.fds-field-input', el => el.value)
      const newToValue = await page.$eval('.to-input-label input.fds-field-input', el => el.value)

      const newFromMatches = newFromValue === expectedFromDateWithZeros || newFromValue === expectedFromDateWithoutZeros
      const newToMatches = newToValue === expectedToDateWithZeros || newToValue === expectedToDateWithoutZeros

      // if (!newFromMatches || !newToMatches) {
      //   throw new Error(`Date values were not set correctly. Expected ${expectedFromDateWithoutZeros} - ${expectedToDateWithoutZeros} or ${expectedFromDateWithZeros} - ${expectedToDateWithZeros}, got ${newFromValue} - ${newToValue}`)
      // }
    }

    return { from: fromValue, to: toValue }
  } catch (error) {
    console.error('Error setting date range:', error)
    throw error
  }
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
    console.log('Navigating to Expedia Partner Central...')
    await page.goto(
      'https://www.expediapartnercentral.com/Account/Logon?signedOff=true',
      {
        waitUntil: ['networkidle0', 'domcontentloaded'],
        timeout: 60000,
      }
    )

    console.log('Waiting for page load...')

    await delay(randomDelay())

    // Wait for email input
    await page.waitForSelector("#emailControl")
    
    // Type email slowly, character by character
    for (let char of email) {
      await page.type("#emailControl", char, { delay: 100 }) // 100ms delay between each character
    }

    // Click continue button
    await page.click("#continueButton")
    
    // Wait before entering password
    console.log('Waiting for password page to load...')

    // Wait for password page to be fully loaded
    try {
      // Wait for the loading indicator to disappear (if any)
      // await page.waitForSelector('.loading-indicator', { 
      //   hidden: true,
      //   timeout: 5000 
      // }).catch(() => console.log('No loading indicator found'))

      // Wait for password field to be visible and ready
      await page.waitForSelector('#passwordControl', {
        visible: true,
        timeout: 10000
      })

      // Additional wait to ensure page is fully loaded
      await delay(4000)

      console.log('Password page loaded, entering password...')
      
      // Type password slowly
      for (let char of password) {
        await page.type('#passwordControl', char, { delay: 100 })
        await delay(50) // Extra small delay between characters
      }

      // Wait a moment before clicking submit
      await delay(5000)
      
      // Click the login button
      await page.click('#signInButton')

    } catch (error) {
      console.log('Error during password entry:', error.message)
      throw error
    }

    // Wait for verification code page using the correct selector
    console.log('Waiting for verification page...')
    await page.waitForSelector('input[name="passcode-input"]', {
      visible: true,
      timeout: 60000,
    })

    // Add delay before fetching verification code
    console.log('Waiting for verification email...')
    await delay(15000) // Wait 15 seconds for email to arrive

    // Get verification code
    const code = await getVerificationCode()
    if (!code) {
      throw new Error('Failed to get verification code from email')
    }
    console.log('Got verification code:', code)

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
    console.log('Clicked the verify button successfully!')

    // Wait for successful login
    await page.waitForNavigation({
      waitUntil: 'networkidle0',
      timeout: 60000,
    })

    console.log('Login successful!')

    // Wait for property table to load
    await page.waitForSelector('.fds-data-table-wrapper', {
      visible: true,
      timeout: 30000
    })

    // Wait for property search input
    await page.waitForSelector('.all-properties__search input.fds-field-input')

    // Get property name from query params
    console.log(`Searching for property: ${propertyName}`)

    // Type property name in search
    await page.type('.all-properties__search input.fds-field-input', propertyName, { delay: 100 })

    // Wait for search results
    await delay(2000)

    // Find and click the property link with more specific selector
    try {
      // Wait for search results to update
      await page.waitForSelector('tbody tr', {
        visible: true,
        timeout: 10000
      })

      // More specific selector for the property link
      const propertySelector = `.property-cell__property-name a[href*="/lodging/home/home"]`
      
      const propertyLink = await page.waitForSelector(propertySelector, {
        visible: true,
        timeout: 10000
      })
      
      if (propertyLink) {
        // Get the text to verify it's the right property
        const linkText = await page.evaluate(el => el.textContent, propertyLink)
        console.log(`Found property: ${linkText}, clicking...`)
        
        try {
          // Click the link and wait for navigation
          await Promise.all([
            page.waitForNavigation({ waitUntil: 'networkidle0', timeout: 30000 }),
            propertyLink.click()
          ])
          
          // Wait for the new page to load
          await delay(8000)

          // Find and click the Reservations link
          console.log('Looking for Reservations link...')
          
          try {
            // Wait for the drawer content to load
            await page.waitForSelector('.uitk-drawer-content', {
              visible: true,
              timeout: 30000
            })

            // Click using JavaScript with the exact structure
            const clicked = await page.evaluate(() => {
              const reservationsItem = Array.from(document.querySelectorAll('.uitk-action-list-item-content'))
                .find(item => {
                  const textDiv = item.querySelector('.uitk-text.overflow-wrap')
                  return textDiv && textDiv.textContent.trim() === 'Reservations'
                })

              if (reservationsItem) {
                const link = reservationsItem.querySelector('a.uitk-action-list-item-link')
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
                timeout: 30000 
              }),
              delay(5000)
            ])

            console.log('Successfully navigated to Reservations page')

            // Wait for date filters to be visible
            console.log('Waiting for date filters...')
            await page.waitForSelector(
              'input[type="radio"][name="dateTypeFilter"]',
              { visible: true, timeout: 30000 }
            )

            // Click the "Checking out" radio button
            console.log('Selecting "Checking out" filter...')
            await page.evaluate(() => {
              const radioButtons = Array.from(document.querySelectorAll('input[type="radio"][name="dateTypeFilter"]'))
              const checkingOutButton = radioButtons.find(radio => 
                radio.parentElement.querySelector('.fds-switch-label').textContent.trim() === 'Checking out'
              )
              if (checkingOutButton) {
                checkingOutButton.click()
              }
            })

            // Wait for radio button click to take effect
            await delay(2000)

            // If start_date and end_date are provided, set the date range
            if (start_date && end_date) {
              await setDateRange(page, start_date, end_date)
              
              // Verify the dates were set correctly
              const dateValues = await page.evaluate(() => {
                return {
                  from: document.querySelector('.from-input-label input.fds-field-input').value,
                  to: document.querySelector('.to-input-label input.fds-field-input').value
                }
              })

              console.log('Set dates:', dateValues)
              
              // Find and click the Apply button
              try {
                // Wait for the apply button to be visible
                await page.waitForSelector('.fds-cell.all-cell-1-4 button.fds-button2.primary', {
                  visible: true,
                  timeout: 10000
                })

                // Click using more specific selector
                await page.evaluate(() => {
                  const applyButton = document.querySelector('.fds-cell.all-cell-1-4 button.fds-button2.primary')
                  if (applyButton) {
                    applyButton.click()
                    return true
                  }
                  throw new Error('Apply button not found')
                })

                console.log('Clicked Apply button, waiting for data to load...')
                
                // Wait for the loading indicator to appear
                await page.waitForSelector('td .fds-loader.is-loading.is-visible', {
                  visible: true,
                  timeout: 10000
                }).catch(() => console.log('Loading indicator did not appear'))

                // Wait for the loading indicator to disappear
                await page.waitForSelector('td .fds-loader.is-loading.is-visible', {
                  hidden: true,
                  timeout: 30000
                })

                console.log('Loading completed, continuing with data processing...')

                // Then continue with your existing code for processing the data...
                console.log('Starting to process reservation data...')

                // Wait for the table to be visible
                await page.waitForSelector('table.fds-data-table', {
                  visible: true,
                  timeout: 30000
                })

                // Wait for data to load and stabilize
                let previousCount = 0
                let attempts = 0
                const maxAttempts = 15 // Increased max attempts

                while (attempts < maxAttempts) {
                  await delay(2000)
                  
                  const currentCount = await page.evaluate(() => {
                    return document.querySelectorAll('td.guestName button.guestNameLink').length
                  })
                  
                  console.log(`Found ${currentCount} reservations on attempt ${attempts + 1}...`)
                  
                  if (currentCount === previousCount && currentCount > 0) {
                    console.log('Data count stabilized')
                    break
                  }
                  
                  previousCount = currentCount
                  attempts++
                }

                // Final verification
                const finalCount = await page.evaluate(() => {
                  return document.querySelectorAll('td.guestName button.guestNameLink').length
                })
                
                console.log(`Final reservation count: ${finalCount}`)
                
                if (finalCount === 0) {
                  throw new Error('No reservations found after multiple attempts')
                }

              } catch (error) {
                console.log('Error with Apply button:', error.message)
                throw error
              }

              // After date range is applied and before scraping data
              console.log('Setting results per page to 100...')
              await page.waitForSelector('.fds-pagination-selector select')
              await page.click('.fds-pagination-selector select')
              await page.select('.fds-pagination-selector select', '100')

              // Wait for data to reload with 100 records
              await delay(3000)
              await page.waitForSelector('table.fds-data-table tbody tr', {
                visible: true,
                timeout: 30000
              })

              // Initialize array for all reservations with Set for tracking duplicates
              const allReservations = []
              const processedReservationIds = new Set()

              // Function to check if there's a next page
              const hasNextPage = async () => {
                return await page.evaluate(() => {
                  const nextButton = document.querySelector('.fds-pagination-button.next button')
                  return nextButton && !nextButton.disabled
                })
              }

              // Function to get total results count
              const getTotalResults = async () => {
                const resultsText = await page.$eval('.fds-pagination-showing-result', el => el.textContent)
                const match = resultsText.match(/of (\d+) Results/)
                return match ? parseInt(match[1]) : 0
              }

              const totalResults = await getTotalResults()
              console.log(`Total reservations to fetch: ${totalResults}`)

              let currentPage = 1
              let hasMore = true

              while (hasMore) {
                try {
                  console.log(`Processing page ${currentPage}...`)
                  
                  // Wait for table data to load
                  await page.waitForSelector('table.fds-data-table tbody tr', {
                    visible: true,
                    timeout: 30000
                  })
                  await delay(2000)

                  // Get reservations from current page
                  const rows = await page.$$('table.fds-data-table tbody tr')
                  
                  for (const row of rows) {
                    try {
                      // Get basic data first
                      const basicData = await page.evaluate(row => {
                        return {
                          guestName: row.querySelector('td.guestName button.guestNameLink span.fds-button2-label')?.textContent.trim() || '',
                          reservationId: row.querySelector('td.reservationId div.fds-cell')?.textContent.trim() || '',
                          confirmationCode: row.querySelector('td.confirmationCode label.confirmationCodeLabel')?.textContent.trim() || '',
                          checkInDate: row.querySelector('td.checkInDate')?.textContent.trim() || '',
                          checkOutDate: row.querySelector('td.checkOutDate')?.textContent.trim() || '',
                          roomType: row.querySelector('td.roomType')?.textContent.trim() || '',
                          bookingAmount: row.querySelector('td.bookingAmount .fds-currency-value')?.textContent.trim() || '',
                          bookedDate: row.querySelector('td.bookedOnDate')?.textContent.trim() || ''
                        }
                      }, row)

                      // Check if we've already processed this reservation
                      if (processedReservationIds.has(basicData.reservationId)) {
                        console.log(`Skipping duplicate reservation: ${basicData.reservationId}`)
                        continue
                      }

                      // Add to processed set
                      processedReservationIds.add(basicData.reservationId)

                      // Get card details
                      const guestNameButton = await row.$('td.guestName button.guestNameLink')
                      await guestNameButton.click()
                      
                      // Wait for initial dialog to appear
                      await page.waitForSelector('.fds-dialog-content', {
                        visible: true,
                        timeout: 8000
                      })

                      // Wait a bit for content to load
                      await delay(2000)

                      // Scroll to the bottom of dialog content and wait
                      await page.evaluate(() => {
                        const dialogContent = document.querySelector('.fds-dialog-content')
                        if (dialogContent) {
                          dialogContent.scrollTo(0, dialogContent.scrollHeight)
                        }
                      })

                      // Wait for content to load after scroll
                      await delay(2000)

                      // Get card details with retry mechanism
                      let cardData = null
                      let retries = 0
                      while (!cardData && retries < 3) {
                        try {
                          cardData = await page.evaluate(() => {
                            const cardNumber = document.querySelector('.cardNumber.replay-conceal bdi')?.textContent.trim() || ''
                            const expiryDate = document.querySelector('.cardDetails .fds-cell.all-cell-1-4.fds-type-color-primary.replay-conceal')?.textContent.trim() || ''
                            const cvv = document.querySelectorAll('.cardDetails .fds-cell.all-cell-1-4.fds-type-color-primary.replay-conceal')[1]?.textContent.trim() || ''
                            
                            return {
                              cardNumber,
                              expiryDate,
                              cvv
                            }
                          })
                        } catch (e) {
                          retries++
                          await delay(1000)
                        }
                      }

                      // Close the side panel
                      try {
                        await page.click('.fds-dialog-header button.dialog-close')
                        await delay(1500)
                      } catch (e) {
                        console.log('Warning: Could not close dialog normally')
                      }

                      // Add to reservations array
                      allReservations.push({
                        ...basicData,
                        ...cardData
                      })

                    } catch (error) {
                      console.log(`Error processing reservation: ${error.message}`)
                      if (basicData) {
                        allReservations.push({
                          ...basicData,
                          cardNumber: 'N/A',
                          expiryDate: 'N/A',
                          cvv: 'N/A'
                        })
                      }
                    }
                  }

                  console.log(`Processed ${allReservations.length} of ${totalResults} reservations`)

                  // Check if there's a next page
                  hasMore = await hasNextPage()
                  if (hasMore) {
                    await page.click('.fds-pagination-button.next button')
                    await delay(2000)
                    currentPage++
                  }

                } catch (pageError) {
                  console.log(`Error processing page ${currentPage}: ${pageError.message}`)
                  // Try to recover by reloading the page
                  await page.reload({ waitUntil: 'networkidle0' })
                  await delay(5000)
                }
              }

              console.log(`Found total ${allReservations.length} reservations`)

              // At the end, verify we have unique reservations
              console.log(`Total unique reservations: ${allReservations.length}`)
              console.log(`Total processed IDs: ${processedReservationIds.size}`)

              // Get current date and time for filename
              const now = new Date()
              const timestamp = now.toISOString().replace(/[:.]/g, '-') // Format: 2024-03-14T10-30-15-000Z

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
                  'CVV'
                ],
                ...allReservations.map(res => [
                  res.guestName,
                  res.reservationId,
                  res.confirmationCode,
                  res.checkInDate,
                  res.checkOutDate,
                  res.roomType,
                  res.bookingAmount,
                  res.bookedDate,
                  res.cardNumber,
                  res.expiryDate,
                  res.cvv
                ])
              ]

              const ws = xlsx.utils.aoa_to_sheet(wsData)
              xlsx.utils.book_append_sheet(workbook, ws, 'Reservations')
              xlsx.writeFile(workbook, `reservations_${timestamp}.xlsx`)
              console.log(`Saved reservation data to reservations_${timestamp}.xlsx`)

              // Close the browser
              // await browser.close()
              return allReservations
            }

            console.log('No reservation data found after multiple retries')
            if (browser) await browser.close()
            return []
          } catch (error) {
            console.log('Error finding/clicking Reservations:', error.message)
            throw error
          }
        } catch (error) {
          console.log(`Error finding/clicking property: ${error.message}`)
          throw error
        }
      } else {
        throw new Error(`Property "${propertyName}" not found`)
      }
    } catch (error) {
      console.log(`Error finding/clicking property: ${error.message}`)
      throw error
    }
  } catch (error) {
    console.error('Error:', error)
    if (browser) await browser.close()
    throw error
  }
}

// Utility function to split date range into 3-day chunks
// function splitDateRange(startDate, endDate) {
//   const chunks = []
//   const start = new Date(startDate)
//   const end = new Date(endDate)

//   let currentStart = new Date(start)
//   while (currentStart < end) {
//     let currentEnd = new Date(currentStart)
//     currentEnd.setDate(currentEnd.getDate() + 2) // Add 2 days to make it a 3-day chunk

//     if (currentEnd > end) {
//       currentEnd = new Date(end)
//     }

//     chunks.push({
//       start: currentStart.toLocaleDateString('en-US'),
//       end: currentEnd.toLocaleDateString('en-US'),
//     })

//     currentStart = new Date(currentEnd)
//     currentStart.setDate(currentStart.getDate() + 1) // Start next chunk from next day
//   }

//   return chunks
// }

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
    res.send('Authentication successful! You can close this window.')
  } catch (error) {
    res.status(500).send('Error retrieving access token: ' + error.message)
  }
})

// Utility function to split date range into 3-day chunks
function splitDateRange(startDate, endDate) {
  const chunks = []
  const start = new Date(startDate)
  const end = new Date(endDate)

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

// Independent API endpoint for Expedia login automation
app.get('/api/expedia', async (req, res) => {
  const { email, password, start_date, end_date, propertyName } = req.query

  if (!email || !password || !start_date || !end_date || !propertyName) {
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

    // Convert date format from MM/DD/YYYY to DD/MM/YYYY
    const convertDateFormat = (dateStr) => {
      const [month, day, year] = dateStr.split('/')
      // Pad day and month with leading zeros if needed
      const paddedDay = day.padStart(2, '0')
      const paddedMonth = month.padStart(2, '0')
      return `${paddedDay}/${paddedMonth}/${year}`
    }

    // Convert input dates to expected format
    const formattedStartDate = convertDateFormat(start_date)
    const formattedEndDate = convertDateFormat(end_date)

    console.log('Original start date:', start_date, '-> Formatted:', formattedStartDate)
    console.log('Original end date:', end_date, '-> Formatted:', formattedEndDate)

    await loginToExpediaPartner(
      email, 
      password, 
      formattedStartDate, 
      formattedEndDate, 
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
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`)
  if (!loadToken()) {
    console.log('Opening browser for authentication...')
    open(`http://localhost:${port}/auth`)
  }
})
