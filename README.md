# vba-webscraper-internsg

Thsi is a simple web scraper made for [intern.sg](https://www.internsg.com/jobs/) using Excel and VBA.

Had an idea to make this while looking for an internship but I didn't want to scroll through all 120 pages. Turns out intern.sg is relatively simple to scrap, as it does not require any user authentication or limit the number of HTTP requests from a given IP address.

I decided to use Excel and VBA for this project as it is the easiest to package and distribute. You can view and edit all the data directly in the same Excel workbook.

Everything you need to run the scraper is self-contained within Web Scraper.xlsm, though the source code is also been included in the src directory for easy reference.

## Usage

![scraper_screenshot](https://user-images.githubusercontent.com/87565927/221873250-bb12db40-c9ad-49ad-9a7b-9848950d7a99.png)

1. Open up Web Scraper.xlsm
2. Click "Click to Run"
2. You can use the "Esc" key to interrupt the macro at any point in time when the program is running

The table will be cleared every time the program is initiated.

Depending on your internet connection, it can take about 10s to scrap 1 page, so you can expect it to run for around 20 min to finish scraping the entire ~120 pages.

Note: Most of the regex patterns and HTML classes are hard coded into the script, and may not continue to work if the website is altered.

## License
[MIT](https://mit-license.org)
