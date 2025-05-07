import scrapy
from datetime import timedelta, datetime
from scrapy.crawler import CrawlerProcess
import scrapy.utils.misc
import scrapy.core.scraper


# This piece of code prevents scrapy spider from raising OSError. I don't know how and why. Found it on GitHub:
# https://github.com/pyinstaller/pyinstaller/issues/4815
def warn_on_generator_with_return_value_stub(spider, callable):
    pass


scrapy.utils.misc.warn_on_generator_with_return_value = warn_on_generator_with_return_value_stub
scrapy.core.scraper.warn_on_generator_with_return_value = warn_on_generator_with_return_value_stub


result = {}


class SingleDateSpider(scrapy.Spider):
    name = "singledate"
    allowed_domains = ["www.bnb.bg"]
    url = ""
    selected_date = None
    current_date = None

    def generate_url(self):
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
        # save the requested date for the result
        self.current_date = self.selected_date
        # get the url for the selected date and parse it
        self.generate_url()
        yield scrapy.Request(self.url, callback=self.parse)

    def parse(self, response):
        # Get the date and the price for USD/BGN exchange rate for the selected date (if available)
        # result_date = response.css("tr.last.center > td:nth-child(1)::text").get()
        result_price = response.css("tr.last.center > td:nth-child(3)::text").get()

        if result_price:
            # add the FX rate to the result dictionary
            result[self.selected_date.strftime('%Y-%m-%d')] = float(result_price.strip())

        else:
            # Generate an url browsing one day in the past and parse it
            self.current_date -= timedelta(days=1)
            self.generate_url()
            yield scrapy.Request(self.url, callback=self.parse)


class RunSpider:
    def __init__(self, dates):
        self.date_list = dates

    def run(self):
        # create the scrapy crawler process
        process = CrawlerProcess({
            'TELNETCONSOLE_ENABLED': False,
            'USER_AGENT': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                          'AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/58.0.3029.110 Safari/537.3',
            'LOG_LEVEL': 'INFO',
            'DOWNLOAD_DELAY': 1,  # 2 seconds between each Request
        })

        # run the scrapy crawler for each date in the "Acquired/Sold" columns in the Revolut PnL file
        for date in self.date_list:
            process.crawl(SingleDateSpider,
                          selected_date=datetime.strptime(date, "%Y-%m-%d")
                          )

        process.start()
