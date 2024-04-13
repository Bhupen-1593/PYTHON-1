import scrapy

class CoffeeScrapy(scrapy.Spider):
    name= 'cof_ext'
    start_urls=['https://www.worldatlas.com/articles/top-10-coffee-consuming-nations.html']
    
    def parse(self,response):
        
        rows=response.xpath('//tbody/tr')
        
        for row in rows:
            rank=row.xpath('.//td[1]/text()').get()
            country=row.xpath('.//td[2]/text()').get()
            consumption=row.xpath('.//td[3]/text()').get()
            
            
            yield{
                'Rank':rank,
                'Country':country,
                'Consumption':consumption
                }
            
            
