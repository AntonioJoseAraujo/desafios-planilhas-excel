import scrapy

class ProxyScraperSpider(scrapy.Spider):
    # identidade
    name = 'proxybot'

    # Request
    def start_requests(self):
        urls = ['https://free-proxy-list.net/web-proxy.html']
        for url in urls:
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        # Montar xpath que retorna linha
        for linha in response.xpath('//table[@class="table table-striped table-bordered"]//tr'):
            yield{
                # Montar individualmente xpath que retorna cada item da linha
                'Proxy name': linha.xpath('./td[1]/a/text()').get(),
                'domain': linha.xpath('./td[2]/text()').get(),
                'country': linha.xpath('./td[3]/text()').get(),
                'speed': linha.xpath('./td[4]/text()').get(),
                'popularity': linha.xpath('./td/div/div/text()').get()
            }

# Ir em pipelines.py onde tem o restante do código
# Em settings.py foi feito algumas alterações