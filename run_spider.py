import scrapy
from datetime import timedelta, datetime
from scrapy.crawler import CrawlerProcess

result = {}


class SingleDateSpider(scrapy.Spider):
    spider_result = result
    name = "singledate"
    allowed_domains = ["www.bnb.bg"]
    url = ""
    selected_date = None
    current_date = None

    def generate_url(self):
        print("Generating url...")
        # Use the selected_date attribute to generate an url to parse
        self.url = (
            f"https://www.bnb.bg/Statistics/StExternalSector/"
            f"StExchangeRates/StERForeignCurrencies/index.htm?"
            f"downloadOper="
            f"&group1=second"
            f"&periodStartDays={str(self.current_date.day)}"
            f"&periodStartMonths={str(self.current_date.month)}"
            f"&periodStartYear={str(self.current_date.year)}"
            f"&periodEndDays={str(self.current_date.day)}"
            f"&periodEndMonths={str(self.current_date.month)}"
            f"&periodEndYear={str(self.current_date.year)}"
            f"&valutes=USD"
            f"&search=true"
            f"&showChart=false"
            f"&showChartButton=true"
        )

    def start_requests(self):
        print("Starting requests...")
        # save the requested date for the result
        self.current_date = self.selected_date
        # get the url for the selected date and parse it
        self.generate_url()
        yield scrapy.Request(self.url, callback=self.parse)

    def parse(self, response):
        print("Parsing...")
        # Get the date and the result_price for USD/BGN exchange rate for the selected date (if available)
        result_date = response.css("td.first.center::text").get()
        result_price = response.css("td.last.center::text").get()

        if result_date:
            print("Parsing successfull. Adding to result...")
            # add the FX rate to the result dictionary
            result[self.selected_date.strftime('%Y-%m-%d')] = float(result_price.strip())
        else:
            print("Parsing failed. Nothing to add...")
            # Generate an url browsing one day in the past and parse it
            self.current_date -= timedelta(days=1)
            self.generate_url()
            yield scrapy.Request(self.url, callback=self.parse)
        result["a"] = 2

class RunSpider:
    def __init__(self, dates):
        self.date_list = dates
        self.result = result

    def run(self):
        print("Running spider...")
        # create the scrapy crawler process
        process = CrawlerProcess({
            'USER_AGENT': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                          'AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/58.0.3029.110 Safari/537.3',
            'LOG_LEVEL': 'INFO',
            'DOWNLOAD_DELAY': 2,  # 2 seconds between each Request
        })

        # run the scrapy crawler for each date in the "Acquired/Sold" columns in the Revolut PnL file
        for date in self.date_list:
            print("Processing spider...")
            process.crawl(SingleDateSpider,
                          selected_date=datetime.strptime(date, "%Y-%m-%d")
                          )

        process.start()
