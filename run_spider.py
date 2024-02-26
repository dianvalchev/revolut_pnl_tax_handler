import scrapy
from datetime import timedelta, datetime
from scrapy.crawler import CrawlerProcess

result = []


class SingleDateSpider(scrapy.Spider):
    name = "singledate"
    allowed_domains = ["www.bnb.bg"]
    url = ""
    selected_date = None
    initial_date = None

    def generate_url(self):

        # Use the selected_date attribute to generate an url to parse
        self.url = (
            f"https://www.bnb.bg/Statistics/StExternalSector/"
            f"StExchangeRates/StERForeignCurrencies/index.htm?"
            f"downloadOper="
            f"&group1=second"
            f"&periodStartDays={str(self.selected_date.day)}"
            f"&periodStartMonths={str(self.selected_date.month)}"
            f"&periodStartYear={str(self.selected_date.year)}"
            f"&periodEndDays={str(self.selected_date.day)}"
            f"&periodEndMonths={str(self.selected_date.month)}"
            f"&periodEndYear={str(self.selected_date.year)}"
            f"&valutes=USD"
            f"&search=true"
            f"&showChart=false"
            f"&showChartButton=true"
        )

    def start_requests(self):
        self.initial_date = self.selected_date
        # Get the url for the selected date and parse it
        self.generate_url()
        yield scrapy.Request(self.url, callback=self.parse)

    def parse(self, response):
        # Get the date and the result_price for USD/BGN exchange rate for the selected date (if available)
        result_date = response.css("td.first.center::text").get()
        result_price = response.css("td.last.center::text").get()

        if result_date:
            # print(f"Exchange rate found for {self.selected_date.strftime('%Y-%m-%d')}...")
            result.append({
                'Date': self.initial_date.strftime('%Y-%m-%d'),
                'Price': float(result_price.strip())
            })
        else:
            # Generate an url browsing one day in the past and parse it

            # temp_current_date = self.selected_date
            self.selected_date -= timedelta(days=1)
            # print(f"No data found for {temp_current_date.strftime('%Y-%m-%d')}. "
            #       f"Moving to {self.selected_date.strftime('%Y-%m-%d')}...")
            self.generate_url()
            yield scrapy.Request(self.url, callback=self.parse)


class RunSpider:
    # Must be in the same file (for now) as the spider to use the global "result" variable
    def __init__(self, dates):
        self.date_list = dates
        self.result = result

    def run(self):
        process = CrawlerProcess({
            'USER_AGENT': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                          'AppleWebKit/537.36 (KHTML, like Gecko) '
                          'Chrome/58.0.3029.110 Safari/537.3',
            'LOG_LEVEL': 'INFO',
            'DOWNLOAD_DELAY': 2,  # 2 seconds between each Request
        })

        for date in self.date_list:
            process.crawl(SingleDateSpider,
                          selected_date=datetime.strptime(date, "%Y-%m-%d")
                          )

        process.start()
