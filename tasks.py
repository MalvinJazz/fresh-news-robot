import logging, os, re, urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from dateutil.relativedelta import relativedelta
from robocorp.tasks import task, get_output_dir
from robocorp import workitems
from RPA.Browser.Selenium import Selenium
from selenium.common.exceptions import ElementClickInterceptedException
from abc import ABC, abstractmethod
from RPA.Excel.Files import Files
from time import sleep

"""Constants for tasks"""
YAHOO_NEWS_SITE = "https://news.yahoo.com/"
LATIMES_SITE = "https://www.latimes.com/"

"""Constants for LATimes"""
PROMO_TITLE = 'promo-title'
PROMO_DESCRIPTION = 'promo-description'
PROMO_TIMESTAMP = 'promo-timestamp'
PROMO_MEDIA = 'promo-media'


class NewInformation():
    """Information about news articles"""
    def __init__(self, title: str, date: datetime, description: str, image_url: str):
        self.__title = title
        self.__date = date
        self.__description = description
        self.__image_url = image_url

    @property
    def title(self):
        return self.__title

    @title.setter
    def title(self, title):
        self.__title = title
    
    @property
    def date(self):
        return self.__date
    
    @date.setter
    def date(self, date):
        self.__date = date

    @property
    def description(self):
        return self.__description

    @description.setter
    def description(self, description):
        self.__description = description

    @property
    def image_url(self):
        return self.__image_url
    
    @image_url.setter
    def image_url(self, image_url):
        self.__image_url = image_url

    def count_of_phrase_occurrences(self, phrase: str):
        return self.__title.count(phrase) + self.__description.count(phrase)

    def has_money(self) -> bool:
        return re.search(r"\$?(\d{1,3},?\d{1,3}\.?\d*)\d* ?(dollars|USD)?", self.__title) is not None or re.search(r"\$?(\d{1,3},?\d{1,3}\.?\d*)\d* ?(dollars|USD)?", self.__description) is not None

class NewsParameters():
    """The parameters to be passed to the robot"""
    __phrase = ""
    __months = 0
    __topic = ""

    def __init__(self) -> None:
        self.items = workitems.inputs.current.payload
        search_parameters = self.items.get("search")
        if search_parameters is not None:
            self.__phrase = search_parameters.get("phrase") if search_parameters.get("phrase") is not None else ""
            self.__months = search_parameters.get("months") if search_parameters.get("months") is not None else 0
            self.__topic = search_parameters.get("topic") if search_parameters.get("topic") is not None else ""

    def get_phrase(self) -> str:
        return self.__phrase
    
    def get_months(self) -> int:
        return self.__months
    
    def get_topic(self) -> str:
        return self.__topic
    
class BaseNewsLogic(ABC):
    """Interface for creating news collection strategies"""

    @abstractmethod
    def open_site(self):
        pass

    @abstractmethod
    def enter_phrase(self, query):
        pass

    @abstractmethod
    def order_and_select_category(self, topic: str):
        pass

    @abstractmethod
    def get_news(self, months: int) -> list[NewInformation]:
        pass
    
    @abstractmethod
    def set_browser(self, browser: Selenium):
        pass

class LATimesLogic(BaseNewsLogic):
    """Search into LATimes site"""
    __site = LATIMES_SITE
    __news_title_class = PROMO_TITLE
    __news_description_class = PROMO_DESCRIPTION
    __news_date_class = PROMO_TIMESTAMP
    __news_image_class = PROMO_MEDIA

    def __init__(self):
        pass

    def open_site(self):
        """Open a news site"""
        self.__browser.open_available_browser(self.__site)

    def enter_phrase(self, query):
        """Enter a phrase to search for"""
        self.__browser.click_button("data:element:search-button")
        self.__browser.input_text("data:element:search-form-input", query)
        self.__browser.click_button("data:element:search-submit-button")
    
    def order_and_select_category(self, topic: str):
        """Order the articles and select category if available"""
        self.__browser.wait_until_page_contains_element("name:s", timedelta(seconds=5))
        self.__browser.select_from_list_by_label("name:s", "Newest")

        clean_topic = topic.strip()
        if clean_topic == "":
            return

        if self.__browser.is_element_visible("class:filters-open-button"):
            self.__browser.click_button("class:filters-open-button")
        
        try:
            self.__browser.wait_until_element_is_visible("//label[@class='checkbox-input-label' and span[text()='{0}']]/input", timeout=timedelta(seconds=5), error=None)
            if self.__browser.is_element_visible("//label[@class='checkbox-input-label' and span[text()='{0}']]/input"):
                self.__browser.select_checkbox("//label[@class='checkbox-input-label' and span[text()='{0}']]/input".format(clean_topic))

            if self.__browser.is_element_visible("class:filters-open-button"):
                self.__browser.click_button_when_visible("class:filters-open-button")
        except AssertionError:
            logging.error(f'Unable to find topic section: {clean_topic}')

    def get_news(self, months: int) -> list[NewInformation]:
        """Collect all articles and its information"""
        locator = "//ul[@class='search-results-module-results-menu']/li"
        search_results = self.__browser.find_elements(locator)

        next_page_link = "//div[@class='search-results-module-next-page']/a"
        max_date_to_search = (datetime.now() - relativedelta(months=months - 1 if months else 0)).replace(day=1)

        news = []
        page = 0
        while self.__browser.is_element_visible(next_page_link):
            for result_index, element in enumerate(search_results, 1):

                timestamp = self.__browser.get_element_attribute(f'{locator}[{result_index}]//p[contains(@class,\'{self.__news_date_class}\')]', "data-timestamp")
                image_url = self.__browser.get_element_attribute(f'{locator}[{result_index}]//div[@class=\'{self.__news_image_class}\']/a/picture/img', "src")

                date_time = datetime.now()
                try:
                    date_time = datetime.fromtimestamp(int(timestamp)/1000.0)
                except ValueError:
                    logging.error(f'Could not convert to date: index({result_index})')
                    pass

                if date_time < max_date_to_search:
                    return news
                
                filename = f'{page}-{timestamp}-{result_index}.png'

                Path(os.path.join(get_output_dir(), "news")).mkdir(parents=True, exist_ok=True)
                urllib.request.urlretrieve(image_url, os.path.join(get_output_dir(), "news", filename))

                new_information = NewInformation(
                    title=self.__browser.get_text(f'{locator}[{result_index}]//h3[@class=\'{self.__news_title_class}\']/a'),
                    description=self.__browser.get_text(f'{locator}[{result_index}]//p[@class=\'{self.__news_description_class}\']'),
                    date=date_time,
                    image_url=filename
                )

                news.append(new_information)

            attempts = 3
            for _ in range(attempts):
                try:
                    self.__browser.click_link(next_page_link) # Successful click, exit the loop
                except ElementClickInterceptedException:
                    sleep(1)

            page = page + 1

        return news


    def set_browser(self, browser: Selenium):
        self.__browser = browser

class SearchContext():
    """Search logic for articles"""

    def __init__(self, browser: Selenium, news_logic: BaseNewsLogic) -> None:
        self.__news_logic = news_logic
        news_logic.set_browser(browser)
    
    def __generate_report(self, results: list[NewInformation],parameters: NewsParameters):
        """Generates excel report for articles collected from selected site"""
        results_dict = {
            'Title':[],
            'Description':[],
            'Date': [],
            'HasMoney': [],
            'Phrase Occurrences': []
        }
        for result in results:
            results_dict['Title'].append(result.title)
            results_dict['Description'].append(result.description)
            results_dict['Date'].append(result.date)
            results_dict['HasMoney'].append(result.has_money())
            results_dict['Phrase Occurrences'].append(result.count_of_phrase_occurrences(parameters.get_phrase()))

        excel = Files()
        wb = excel.create_workbook()
        wb.create_worksheet('Results')
        excel.append_rows_to_worksheet(results_dict,header=True,name='Results')
        wb.save(os.path.join(get_output_dir(),'fresh_news.xlsx'))

    def search(self, parameters: NewsParameters):
        """Chain of execution for collect and generate report of results"""
        self.__news_logic.open_site()
        self.__news_logic.enter_phrase(parameters.get_phrase())
        self.__news_logic.order_and_select_category(parameters.get_topic())
        results = self.__news_logic.get_news(parameters.get_months())
        self.__generate_report(results, parameters)

    @property
    def news_logic(self):
        return self.__news_logic
    
    @news_logic.setter
    def news_logic(self, news_logic: BaseNewsLogic) -> None:
        self.__news_logic = news_logic

@task
def get_fresh_news_task():
    """Robot to get fresh news by month and generate a report"""

    search_context = SearchContext(Selenium(), LATimesLogic())
    search_context.search(NewsParameters())
    