import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time
import logging
import os
import json
import traceback
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import sys
import re
import urllib.parse
import cloudscraper
import requests
from fake_useragent import UserAgent
import tls_client

# Configure logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Create file handler
file_handler = logging.FileHandler('indeed_scrape_log.txt', mode='w')
file_handler.setLevel(logging.DEBUG)

# Create console handler
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)

# Create formatter
formatter = logging.Formatter('%(asctime)s - %(levelname)s: %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add handlers to logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

class IndeedScraper:
    def __init__(self, chromedriver_path=None, existing_browser_port=None):
        """
        Initialize Indeed Scraper with option to use existing browser session
        
        :param chromedriver_path: Optional path to ChromeDriver executable
        :param existing_browser_port: Port of an existing Chrome browser debugging session
        """
        self.driver = None
        self.ua = UserAgent()
        
        try:
            # Initialize TLS client for advanced request handling
            self.tls_session = tls_client.Session(
                client_identifier="chrome_110",
                random_tls_extension_order=True
            )
            
            # CloudScraper for additional bot bypassing
            self.cloudscaper = cloudscraper.create_scraper(
                browser={
                    'browser': 'chrome',
                    'platform': 'windows',
                    'mobile': False
                }
            )
            
            # Job titles and URLs to search
            self.job_searches = [
                {
                    'title': 'ui/ux designer', 
                    'url': 'https://malaysia.indeed.com/jobs?q=ui%2Fux+designer&l=Malaysia&from=searchOnHP&vjk=c365239bc9243b1f'
                },
                {
                    'title': 'ux designer', 
                    'url': 'https://malaysia.indeed.com/jobs?q=ux+designer&l=Malaysia&from=searchOnDesktopSerp&vjk=91129577b6bfc8ae'
                },
                {
                    'title': 'ux design lead', 
                    'url': 'https://malaysia.indeed.com/jobs?q=ux+design+lead&l=Malaysia&from=searchOnDesktopSerp&vjk=b550e0f933e6417b'
                },
                {
                    'title': 'ux design manager', 
                    'url': 'https://malaysia.indeed.com/jobs?q=ux+design+manager&l=Malaysia&from=searchOnDesktopSerp&vjk=c365239bc9243b1f'
                },
                {
                    'title': 'product design lead', 
                    'url': 'https://malaysia.indeed.com/jobs?q=product+design+lead&l=Malaysia&from=searchOnDesktopSerp&vjk=b550e0f933e6417b'
                },
                {
                    'title': 'product design manager', 
                    'url': 'https://malaysia.indeed.com/jobs?q=product+design+manager&l=Malaysia&from=searchOnDesktopSerp&vjk=c365239bc9243b1f'
                }
            ]
            
            # Attach to existing browser or create new session
            if existing_browser_port:
                # Connect to existing Chrome browser session
                logger.info(f"Connecting to existing Chrome browser on port {existing_browser_port}")
                chrome_options = uc.ChromeOptions()
                chrome_options.debugger_address = f'127.0.0.1:{existing_browser_port}'
                self.driver = uc.Chrome(options=chrome_options)
            else:
                # Setup Chrome options with advanced bot evasion
                chrome_options = uc.ChromeOptions()
                chrome_options.add_argument("--start-maximized")
                chrome_options.add_argument(f"user-agent={self.ua.random}")
                chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
                chrome_options.add_experimental_option('useAutomationExtension', False)
                
                # Initialize the driver with retry mechanism
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        logger.info(f"Initializing ChromeDriver (Attempt {attempt + 1}/{max_retries})...")
                        self.driver = uc.Chrome(options=chrome_options)
                        
                        # Additional browser fingerprint randomization
                        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                            "source": """
                            Object.defineProperty(navigator, 'webdriver', {
                                get: () => undefined
                            })
                            """
                        })
                        
                        break  # Success, exit retry loop
                    except Exception as init_error:
                        logger.warning(f"ChromeDriver initialization failed: {init_error}")
                        if attempt == max_retries - 1:
                            raise
                        time.sleep(2)  # Wait before retry
            
            # Prepare output directory
            self.output_dir = 'indeed_output'
            os.makedirs(self.output_dir, exist_ok=True)
        
        except Exception as e:
            logger.error(f"Indeed Scraper Initialization Error: {e}")
            logger.error(traceback.format_exc())
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
            raise

    def cloudflare_bypass(self, url):
        """
        Advanced Cloudflare bypass technique
        
        :param url: URL to bypass
        :return: Requests session with bypassed Cloudflare
        """
        try:
            # Multiple bypass strategies
            bypass_strategies = [
                # Strategy 1: CloudScraper
                lambda: self.cloudscaper.get(url),
                
                # Strategy 2: TLS Client
                lambda: self.tls_session.get(url, 
                    headers={
                        'User-Agent': self.ua.random,
                        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
                        'Accept-Language': 'en-US,en;q=0.5',
                        'Referer': 'https://www.google.com/',
                        'DNT': '1',
                        'Connection': 'keep-alive',
                        'Upgrade-Insecure-Requests': '1'
                    }
                ),
                
                # Strategy 3: Requests with custom headers
                lambda: requests.get(url, 
                    headers={
                        'User-Agent': self.ua.random,
                        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
                        'Accept-Language': 'en-US,en;q=0.5',
                        'Referer': 'https://www.google.com/',
                        'DNT': '1',
                        'Connection': 'keep-alive',
                        'Upgrade-Insecure-Requests': '1'
                    }
                )
            ]
            
            # Try each bypass strategy
            for strategy in bypass_strategies:
                try:
                    response = strategy()
                    
                    # Check if bypass was successful
                    if response.status_code == 200:
                        logger.info(f"Successful Cloudflare bypass for {url}")
                        return response
                except Exception as strategy_error:
                    logger.warning(f"Bypass strategy failed: {strategy_error}")
            
            # If all strategies fail
            logger.error(f"Failed to bypass Cloudflare for {url}")
            return None
        
        except Exception as e:
            logger.error(f"Cloudflare bypass error: {e}")
            return None

    def diagnose_blocking(self):
        """
        Comprehensive diagnosis of potential blocking mechanisms
        
        :return: Dictionary of diagnostic information
        """
        diagnostics = {
            'page_source_length': 0,
            'status_code': None,
            'blocking_indicators': [],
            'headers': {},
            'cookies': {}
        }
        
        try:
            # Page source length
            diagnostics['page_source_length'] = len(self.driver.page_source)
            
            # Check for common blocking indicators in page source
            blocking_patterns = [
                'captcha', 
                'robot', 
                'verification', 
                'suspicious activity', 
                'blocked', 
                'challenge'
            ]
            
            page_source = self.driver.page_source.lower()
            diagnostics['blocking_indicators'] = [
                pattern for pattern in blocking_patterns 
                if pattern in page_source
            ]
            
            # Collect browser headers
            try:
                # Execute JavaScript to get navigator properties
                diagnostics['browser_info'] = self.driver.execute_script("""
                return {
                    'userAgent': navigator.userAgent,
                    'platform': navigator.platform,
                    'language': navigator.language,
                    'webdriver': navigator.webdriver,
                    'hardwareConcurrency': navigator.hardwareConcurrency
                };
                """)
            except Exception as header_error:
                logger.warning(f"Could not collect browser headers: {header_error}")
            
            # Collect cookies
            try:
                diagnostics['cookies'] = {
                    cookie['name']: cookie['value'] 
                    for cookie in self.driver.get_cookies()
                }
            except Exception as cookie_error:
                logger.warning(f"Could not collect cookies: {cookie_error}")
        
        except Exception as e:
            logger.error(f"Diagnostic error: {e}")
        
        return diagnostics

    def manual_verification(self, job_title, search_url):
        """
        Robust manual verification with error recovery
        
        :param job_title: Current job title being searched
        :param search_url: URL to navigate to
        :return: Boolean indicating whether to proceed
        """
        try:
            # Reinitialize driver if it's no longer active
            if not self.is_driver_active():
                logger.warning("Driver is not active. Reinitializing...")
                self.__init__()
            
            # Navigate to the search URL
            logger.info(f"Navigating to {search_url}")
            self.driver.get(search_url)
            
            # Extended wait for page load and potential challenges
            time.sleep(10)
            
            # Create a detailed human verification popup
            root = tk.Tk()
            root.title("Human Verification Required")
            root.geometry("700x600")
            
            # Verification instructions
            verification_steps = [
                "1. Browser is fully open and maximized",
                "2. Scroll through the page naturally",
                "3. Move mouse cursor randomly",
                "4. Click on some job listings",
                "5. Hover over different page elements",
                "6. Wait for a few moments between actions",
                "7. Ensure no CAPTCHA or verification blocks remain"
            ]
            
            # Main frame
            main_frame = tk.Frame(root)
            main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
            
            # Title
            title_label = tk.Label(main_frame, text="Human Interaction Verification", font=('Arial', 14, 'bold'))
            title_label.pack(pady=10)
            
            # Instructions
            instructions_frame = tk.Frame(main_frame)
            instructions_frame.pack(pady=10)
            
            for step in verification_steps:
                step_label = tk.Label(instructions_frame, text=step, anchor='w', justify=tk.LEFT)
                step_label.pack(fill=tk.X)
            
            # Verification status
            status_var = tk.StringVar(value="Not Verified")
            status_label = tk.Label(main_frame, textvariable=status_var, font=('Arial', 12))
            status_label.pack(pady=10)
            
            # Verification variables
            verification_complete = False
            
            def update_status(message):
                status_var.set(message)
                root.update()
            
            def on_verify():
                nonlocal verification_complete
                verification_complete = True
                update_status("Verification Successful!")
                root.destroy()
            
            def on_cancel():
                nonlocal verification_complete
                verification_complete = False
                update_status("Verification Cancelled")
                root.destroy()
            
            # Verification buttons
            button_frame = tk.Frame(main_frame)
            button_frame.pack(pady=10)
            
            verify_btn = tk.Button(button_frame, text="I've Completed Human Interaction", command=on_verify)
            verify_btn.pack(side=tk.LEFT, padx=10)
            
            cancel_btn = tk.Button(button_frame, text="Cancel", command=on_cancel)
            cancel_btn.pack(side=tk.LEFT, padx=10)
            
            # Provide guidance text
            guidance_text = tk.Text(main_frame, height=5, width=80, wrap=tk.WORD)
            guidance_text.insert(tk.END, "Guidance:\n")
            guidance_text.insert(tk.END, "- Simulate natural browsing behavior\n")
            guidance_text.insert(tk.END, "- Move slowly and randomly\n")
            guidance_text.insert(tk.END, "- Interact with page elements genuinely\n")
            guidance_text.config(state=tk.DISABLED)
            guidance_text.pack(pady=10)
            
            # Keep the window open and wait for user action
            root.mainloop()
            
            # If verification is not complete, return False
            if not verification_complete:
                logger.warning(f"Manual verification failed for {job_title}")
                return False
            
            # Additional human-like interactions
            try:
                # Scroll page
                self.driver.execute_script("window.scrollBy(0, 250);")
                time.sleep(1)
                self.driver.execute_script("window.scrollBy(0, -100);")
                time.sleep(1)
                
                # Hover over some elements
                job_cards = self.driver.find_elements(By.CSS_SELECTOR, 'div.job_seen_beacon')
                if job_cards:
                    webdriver.ActionChains(self.driver).move_to_element(job_cards[0]).perform()
                    time.sleep(1)
            except Exception as interaction_error:
                logger.warning(f"Additional human interaction error: {interaction_error}")
            
            logger.info(f"Manual verification successful for {job_title}")
            return True
        
        except Exception as e:
            logger.error(f"Critical error during manual verification for {job_title}: {e}")
            logger.error(traceback.format_exc())
            return False
    
    def is_driver_active(self):
        """
        Check if the WebDriver is still active and usable
        
        :return: Boolean indicating driver's active status
        """
        try:
            # Check if we can execute a simple script
            self.driver.execute_script("return 1;")
            return True
        except Exception:
            return False

    def scroll_and_load_jobs(self, max_attempts=10):
        """
        Advanced scrolling and job loading mechanism
        
        :param max_attempts: Maximum number of attempts to load more jobs
        :return: Number of new jobs loaded
        """
        initial_job_count = len(self.driver.find_elements(By.CSS_SELECTOR, 'div.job_seen_beacon'))
        
        for attempt in range(max_attempts):
            try:
                # Scroll to the bottom of the page
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)  # Wait for potential dynamic content
                
                # Try to click "Next" button if exists
                try:
                    next_button = self.driver.find_element(By.CSS_SELECTOR, 'a[data-testid="pagination-page-next"]')
                    if 'disabled' not in next_button.get_attribute('class'):
                        next_button.click()
                        time.sleep(2)
                except NoSuchElementException:
                    logger.info("No more pages to load")
                    break
                
                # Check if new jobs were loaded
                current_job_count = len(self.driver.find_elements(By.CSS_SELECTOR, 'div.job_seen_beacon'))
                if current_job_count > initial_job_count:
                    logger.info(f"Loaded more jobs: {current_job_count} total")
                    initial_job_count = current_job_count
                else:
                    logger.info("No new jobs loaded in this attempt")
                    break
            
            except Exception as e:
                logger.warning(f"Error during job loading (Attempt {attempt + 1}): {e}")
                time.sleep(2)
        
        return initial_job_count

    def extract_job_details(self, job_title):
        """
        Extract comprehensive job details from current page
        
        :param job_title: Job title being searched (for keyword tracking)
        :return: List of job dictionaries with enhanced details
        """
        jobs = []
        try:
            job_cards = self.driver.find_elements(By.CSS_SELECTOR, 'div.job_seen_beacon')
            
            for card in job_cards:
                try:
                    # Extract basic details
                    title_elem = card.find_element(By.CSS_SELECTOR, 'h2.jobTitle span[title]')
                    company_elem = card.find_element(By.CSS_SELECTOR, 'span.companyName')
                    location_elem = card.find_element(By.CSS_SELECTOR, 'div.companyLocation')
                    link_elem = card.find_element(By.CSS_SELECTOR, 'h2.jobTitle a')
                    
                    # Try to extract salary (not always present)
                    try:
                        salary_elem = card.find_element(By.CSS_SELECTOR, 'div.metadata.salary-info-container')
                        salary = salary_elem.text.strip()
                    except NoSuchElementException:
                        salary = 'Not specified'
                    
                    # Prepare job dictionary
                    job = {
                        'platform': 'Indeed Malaysia',
                        'job_title': title_elem.text.strip(),
                        'company_name': company_elem.text.strip(),
                        'location': location_elem.text.strip(),
                        'salary_range': salary,
                        'link': link_elem.get_attribute('href'),
                        'search_keywords': job_title
                        
                        }
                    
                    jobs.append(job)
                
                except Exception as detail_error:
                    logger.warning(f"Could not extract job details: {detail_error}")
        
        except Exception as e:
            logger.error(f"Error extracting job details: {e}")
        
        return jobs

    def scrape_job_listings(self):
        """
        Main scraping method for Indeed job listings
        
        :return: List of all scraped jobs
        """
        all_jobs = []
        
        for job_search in self.job_searches:
            job_title = job_search['title']
            search_url = job_search['url']
            
            try:
                # Manual verification with option to proceed or skip
                if not self.manual_verification(job_title, search_url):
                    logger.warning(f"Skipping {job_title} as per user request")
                    continue
                
                # Scroll and load jobs
                self.scroll_and_load_jobs()
                
                # Extract job details
                jobs = self.extract_job_details(job_title)
                
                all_jobs.extend(jobs)
                logger.info(f"Scraped {len(jobs)} jobs for {job_title}")
            
            except Exception as e:
                logger.error(f"Error scraping {job_title}: {e}")
                logger.error(traceback.format_exc())
        
        return all_jobs

    def save_to_excel(self, jobs):
        """
        Save scraped jobs to Excel with comprehensive columns
        
        :param jobs: List of job dictionaries
        """
        try:
            df = pd.DataFrame(jobs)
            
            # Reorder columns for better readability
            column_order = [
                'platform', 
                'job_title', 
                'company_name', 
                'location', 
                'salary_range', 
                'link', 
                'search_keywords'
            ]
            
            # Ensure all columns exist, fill with empty string if not
            for col in column_order:
                if col not in df.columns:
                    df[col] = ''
            
            # Select and reorder columns
            df = df[column_order]
            
            # Save to Excel
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(self.output_dir, f'indeed_jobs_{timestamp}.xlsx')
            df.to_excel(filename, index=False)
            logger.info(f"Jobs saved to {filename}")
            return filename
        
        except Exception as e:
            logger.error(f"Error saving to Excel: {e}")
            return None

def main():
    """
    Main function to run the Indeed scraper
    Allows connecting to an existing Chrome browser session
    """
    # Default port for Chrome remote debugging (you can change this)
    existing_browser_port = 9222
    
    # Create scraper instance with existing browser session
    scraper = IndeedScraper(existing_browser_port=existing_browser_port)
    
    try:
        # Scrape job listings
        all_jobs = scraper.scrape_job_listings()
        
        # Save to Excel
        output_file = scraper.save_to_excel(all_jobs)
        
        logger.info(f"Scraping completed. Jobs saved to {output_file}")
    
    except Exception as e:
        logger.error(f"Scraping failed: {e}")
    
    finally:
        # Close the browser
        if scraper.driver:
            scraper.driver.quit()

if __name__ == "__main__":
    main()
