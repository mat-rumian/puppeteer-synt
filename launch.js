const puppeteer = require('puppeteer');
const SumoLogger = require('sumo-logger');

const USER_NAME = 'test-somestfuf-xyz@outlook.com';
const PASSWORD = '!Asdqwe12';

const URL = 'https://www.office.com/';

const opts = {
    endpoint: process.env.SUMO_LOGS_ENDPOINT_URL,
    clientUrl: URL,
    sourceCategory: 'synthetic/puppeteer/o365',
    returnPromise: false,
};

// Instantiate the SumoLogger
const sumoLogger = new SumoLogger(opts);

const SELECTORS = {
    'username_input': '#i0116',
    'password_input': '#i0118',
    'login_btn': `#hero-banner-sign-in-to-office-365-link`,
    'sign_in_next_btn': '#idSIButton9',
    'sign_in_btn': '#idSIButton9',
    'stay_signed_in_checkbox': '#KmsiCheckboxField',
    'stay_signed_in_yes_btn': '#idSIButton9',
    'onedrive_icon': '#ShellDocuments_link > div > i'
};

(async () => {
    async function clickAndSetFieldValue(selectorName, selectors, value, del) {
        console.log(`Setting value for: ${selectorName}`);
        await page.waitForSelector(selectors[selectorName], {timeout: del * 1000});
        await page.type(selectors[selectorName], String(value), );
    };

    async function click(selectorName, selectors, del)  {
        console.log(`Clicking on: ${selectorName}`);
        await page.waitForSelector(selectors[selectorName], {timeout: del * 1000});
        await page.click(selectors[selectorName]);
    };

    async function getPerformanceTimings() {
        let pageTitle = await page.title()
        const performanceTimingJson = await page.evaluate(() => JSON.stringify(window.performance.timing))
        const performanceTiming = JSON.parse(performanceTimingJson)
        const startToInteractive = performanceTiming.domInteractive - performanceTiming.navigationStart

        let data = {
            'pageTitle': pageTitle,
            'rawPerformanceTiming': performanceTiming,
            'navigationToInteractiveMs': startToInteractive,
        }
        console.log(`Navigation to interactive "${pageTitle}" took: ${startToInteractive} ms`)
        console.log(data)
    };

    // Launch Browser and navigate to Login Page
    let browser = await puppeteer.launch({
        headless: false,
        slowMo: 1500,
        executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    })
    let page = await browser.newPage()
    await page.goto(URL)

    // Main Page
    await getPerformanceTimings()
    // Click Login
    await click('login_btn', SELECTORS,5)

    //// Login Username Page
    await getPerformanceTimings()
    //Set Username
    await clickAndSetFieldValue('username_input', SELECTORS, USER_NAME, 5)
    // Click Next
    await click('sign_in_next_btn', SELECTORS,5)

    //// Login Password Page
    await getPerformanceTimings()
    // Set Password
    await clickAndSetFieldValue('password_input', SELECTORS, PASSWORD, 5)
    // Click Sign In
    await click('sign_in_btn', SELECTORS,5)

    //// Stay Signed In Page
    await getPerformanceTimings()
    // Click Stay Signed In Checkbox
    await click('stay_signed_in_checkbox', SELECTORS,5)
    // Click Stay Signed In Yes
    await click('stay_signed_in_yes_btn', SELECTORS,5)

    //// Office App Page
    await getPerformanceTimings()
    // Click Stay Signed In Yes
    await click('onedrive_icon', SELECTORS,5)

    //// OneDrive App Page
    await getPerformanceTimings()

    await sumoLogger.flushLogs();

    await page.close();
    await browser.close();
})()
