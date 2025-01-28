import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
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

# Configure logging
import logging
import sys

# Create logger
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)

# Create file handler
file_handler = logging.FileHandler('glassdoor_scrape_log.txt', mode='w')
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

class GlassdoorScraper:
    def __init__(self, chromedriver_path):
        """
        Initialize Glassdoor Scraper
        
        :param chromedriver_path: Path to ChromeDriver executable
        """
        self.driver = None
        try:
            # Validate ChromeDriver path
            if not os.path.exists(chromedriver_path):
                raise FileNotFoundError(f"ChromeDriver not found at {chromedriver_path}")
            
            # Job titles to search
            self.job_titles = [
                'ui/ux designer', 
                'ux designer', 
                'ux design lead', 
                'ux design manager', 
                'product design lead', 
                'product design manager'
            ]
            
            # Setup Chrome options
            chrome_options = uc.ChromeOptions()
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--disable-extensions")
            
            # Set ChromeDriver service
            service = Service(chromedriver_path)
            
            # Initialize the driver
            logger.info("Initializing ChromeDriver for Glassdoor...")
            self.driver = uc.Chrome(
                service=service, 
                options=chrome_options,
                version_main=131
            )
            logger.info("ChromeDriver initialized successfully")
            
            # Prepare output directory
            self.output_dir = 'glassdoor_output'
            os.makedirs(self.output_dir, exist_ok=True)
        
        except Exception as e:
            logger.error(f"Glassdoor Scraper Initialization Error: {e}")
            logger.error(traceback.format_exc())
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
            raise

    def manual_verification(self, job_title):
        """
        Manual verification popup for user to handle robot detection
        
        :param job_title: Current job title being searched
        :return: Boolean indicating whether to proceed
        """
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        result = messagebox.askyesno(
            "Manual Verification Required", 
            f"Scraping Glassdoor for '{job_title}'.\n\n"
            "Please:\n"
            "1. Open the browser\n"
            "2. Navigate to Glassdoor\n"
            "3. Search for the job title\n"
            "4. Handle any CAPTCHA or verification\n\n"
            "Click 'Yes' when ready to continue\n"
            "Click 'No' to skip this job title"
        )
        
        root.destroy()
        return result

    def manual_popup_handler(self, job_title):
        """
        Manual intervention method for handling popups during scraping
        
        :param job_title: Current job title being scraped
        :return: Boolean indicating whether to continue scraping
        """
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        # Create a popup message
        result = messagebox.askyesno(
            "Popup Detected", 
            f"A popup is blocking the scraping of {job_title} jobs.\n\n"
            "Have you manually closed the popup?\n\n"
            "Click 'Yes' to continue scraping, 'No' to skip this job title."
        )
        
        root.destroy()
        return result

    def sort_jobs(self, sort_type='recent'):
        """
        Sort job listings by most recent or most relevant
        
        :param sort_type: 'recent' or 'relevant'
        :return: Boolean indicating if sorting was successful
        """
        try:
            # Wait for sort button to be clickable
            sort_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-test="sortBy"]'))
            )
            sort_button.click()
            
            # Wait for dropdown to appear
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.Pill_bottomSheetContainer__pUKiw[data-test="sortBy-dropdown"]'))
            )
            
            # Select sorting option
            if sort_type == 'recent':
                sort_xpath = "//button[contains(text(), 'Most recent')]"
            else:
                sort_xpath = "//button[contains(text(), 'Most relevant')]"
            
            # Find and click the sorting option
            recent_option = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, sort_xpath))
            )
            recent_option.click()
            
            # Wait for page to reload with new sorting
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'li.JobsList_jobListItem__wjTHv[data-test="jobListing"]'))
            )
            
            return True
        except Exception as e:
            logger.error(f"Error sorting jobs: {e}")
            return False

    def scroll_and_load_jobs(self, max_attempts=10):
        """
        Advanced scrolling and job loading mechanism
        
        :param max_attempts: Maximum number of attempts to load more jobs
        :return: Number of new jobs loaded
        """
        initial_job_count = len(self.driver.find_elements(By.CSS_SELECTOR, 'li.JobsList_jobListItem__wjTHv[data-test="jobListing"]'))
        
        for attempt in range(max_attempts):
            try:
                # Scroll to bottom
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                
                # Try to click "Show more" button if exists
                try:
                    load_more_button = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-test="load-more"]'))
                    )
                    load_more_button.click()
                    time.sleep(2)
                except Exception:
                    pass
                
                # Try JavaScript click if regular click fails
                try:
                    self.driver.execute_script("""
                        var loadMoreButtons = document.querySelectorAll('button[data-test="load-more"]');
                        if (loadMoreButtons.length > 0) {
                            loadMoreButtons[loadMoreButtons.length - 1].click();
                        }
                    """)
                    time.sleep(2)
                except Exception:
                    pass
                
                # Check if new jobs were loaded
                current_job_count = len(self.driver.find_elements(By.CSS_SELECTOR, 'li.JobsList_jobListItem__wjTHv[data-test="jobListing"]'))
                
                if current_job_count > initial_job_count:
                    logger.info(f"Loaded more jobs: {current_job_count - initial_job_count} new jobs")
                    initial_job_count = current_job_count
                else:
                    logger.info("No new jobs loaded in this attempt")
                    break
            
            except Exception as e:
                logger.warning(f"Error in job loading attempt {attempt + 1}: {e}")
        
        return initial_job_count

    def close_popups(self):
        """
        Close various popup windows that might interrupt scraping
        
        :return: Number of popups closed
        """
        popups_closed = 0
        
        # List of potential popup selectors
        popup_selectors = [
            # Close button for authentication modal
            'button.CloseButton[type="button"] svg.CloseIcon',
            # Alternative close button
            'button.CloseButton[type="button"]',
            # General modal close buttons
            'div.ContentSection button.CloseButton',
            # Additional potential close buttons
            '[data-test="authModalContainerV2-content"] button.CloseButton'
        ]
        
        for selector in popup_selectors:
            try:
                # Find all matching close buttons
                close_buttons = self.driver.find_elements(By.CSS_SELECTOR, selector)
                
                for button in close_buttons:
                    try:
                        # Wait for button to be clickable
                        WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                        )
                        
                        # Try different methods to click
                        try:
                            # Try regular click
                            button.click()
                            popups_closed += 1
                            logger.info(f"Closed popup with selector: {selector}")
                        except Exception:
                            # If regular click fails, try JavaScript click
                            try:
                                self.driver.execute_script("arguments[0].click();", button)
                                popups_closed += 1
                                logger.info(f"Closed popup using JavaScript with selector: {selector}")
                            except Exception as js_err:
                                logger.warning(f"Could not close popup with selector {selector}: {js_err}")
                        
                        # Short wait after closing
                        time.sleep(1)
                    
                    except Exception as button_err:
                        logger.warning(f"Error processing close button: {button_err}")
            
            except Exception as selector_err:
                logger.warning(f"Error finding popup with selector {selector}: {selector_err}")
        
        return popups_closed

    def verify_job_count(self, expected_count, actual_count):
        """
        Verify the job count and log discrepancies
        
        :param expected_count: Initially estimated job count
        :param actual_count: Actual number of jobs scraped
        :return: Boolean indicating if count is acceptable
        """
        # Allow for 10% variance in job count
        variance_threshold = 0.1
        
        if expected_count == float('inf'):
            return True
        
        lower_bound = expected_count * (1 - variance_threshold)
        upper_bound = expected_count * (1 + variance_threshold)
        
        if lower_bound <= actual_count <= upper_bound:
            logger.info(f"Job count verification passed. Expected: {expected_count}, Actual: {actual_count}")
            return True
        else:
            logger.warning(
                f"Job count discrepancy detected. "
                f"Expected: {expected_count}, Actual: {actual_count}, "
                f"Variance: {abs(actual_count - expected_count) / expected_count * 100:.2f}%"
            )
            return False

    def scroll_and_load_comprehensive(self, max_attempts=30):
        """
        Enhanced scrolling method to load more jobs comprehensively
        
        :param max_attempts: Maximum number of scroll attempts
        :return: Number of new jobs loaded
        """
        initial_job_count = len(self.driver.find_elements(By.CSS_SELECTOR, 'li.JobsList_jobListItem__wjTHv[data-test="jobListing"]'))
        
        # Different scrolling strategies
        scroll_strategies = [
            # Scroll to bottom of page
            lambda: self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);"),
            
            # Scroll through job list
            lambda: self.driver.execute_script("""
                let jobList = document.querySelector('ul.JobsList_jobList__Qltsb');
                if (jobList) { jobList.scrollTop = jobList.scrollHeight; }
            """),
            
            # Click "Show More" if available
            lambda: self.click_show_more_jobs()
        ]
        
        new_jobs_loaded = 0
        for attempt in range(max_attempts):
            # Try each scrolling strategy
            for strategy in scroll_strategies:
                try:
                    strategy()
                    time.sleep(2)  # Wait for potential dynamic loading
                except Exception as e:
                    logger.debug(f"Scroll strategy failed: {e}")
            
            # Check for new jobs
            current_job_count = len(self.driver.find_elements(By.CSS_SELECTOR, 'li.JobsList_jobListItem__wjTHv[data-test="jobListing"]'))
            
            if current_job_count > initial_job_count:
                new_jobs_loaded = current_job_count - initial_job_count
                logger.info(f"Loaded {new_jobs_loaded} new jobs in attempt {attempt + 1}")
                initial_job_count = current_job_count
            else:
                # If no new jobs, try clicking "Show More" explicitly
                try:
                    show_more_button = self.driver.find_element(By.CSS_SELECTOR, 'button[data-test="load-more"]')
                    show_more_button.click()
                    time.sleep(2)
                except Exception:
                    # No more jobs to load
                    break
        
        return new_jobs_loaded

    def scrape_jobs(self):
        """
        Scrape job listings from Glassdoor for multiple job titles in Malaysia
        
        :return: List of job dictionaries with title, URL, and other details
        """
        # Comprehensive URLs for Malaysia job searches
        job_search_urls = {
            'ui/ux designer': [
                'https://www.glassdoor.com/Job/jobs.htm?sc.occupationParam=+ui%2Fux+designer&sc.locationSeoString=Malaysia&locId=170&locT=N'
            ],
            'ux designer': [
                'https://www.glassdoor.com/Job/malaysia-ux-designer-jobs-SRCH_IL.0,8_IN170_KO9,20.htm'
            ],
            'ux design lead': [
                'https://www.glassdoor.com/Job/malaysia--ux-design-lead-jobs-SRCH_IL.0,8_IN170_KO9,24.htm'
            ],
            'ux design manager': [
                'https://www.glassdoor.com/Job/malaysia-ux-design-manager-jobs-SRCH_IL.0,8_IN170_KO9,26.htm'
            ],
            'product design lead': [
                'https://www.glassdoor.com/Job/malaysia-product-design-lead-jobs-SRCH_IL.0,8_IN170_KO9,28.htm'
            ],
            'product design manager': [
                'https://www.glassdoor.com/Job/malaysia-product-design-manager-jobs-SRCH_IL.0,8_IN170_KO9,31.htm'
            ]
        }

        all_jobs = []
        job_summary = {}  # To track job counts per search
        unique_job_urls = set()

        for job_title, urls in job_search_urls.items():
            for search_url in urls:
                try:
                    # Navigate to search URL
                    self.driver.get(search_url)
                    
                    # Close any popups first
                    popup_attempts = 0
                    while True:
                        popups_closed = self.close_popups()
                        if popups_closed == 0:
                            break
                        popup_attempts += 1
                        if popup_attempts > 5:
                            # If automatic closing fails multiple times, ask for manual intervention
                            if not self.manual_popup_handler(job_title):
                                logger.warning(f"Skipping {job_title} due to persistent popup")
                                break
                    
                    # Sort jobs by most recent first
                    if not self.sort_jobs('recent'):
                        logger.warning(f"Could not sort jobs by most recent for {job_title}")
                    
                    # Wait for job listings to load
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'li.JobsList_jobListItem__wjTHv[data-test="jobListing"]'))
                    )
                    
                    # Try to extract total job count
                    try:
                        job_count_elem = self.driver.find_element(By.CSS_SELECTOR, 'h1.SearchResultsHeader_jobCount__eHngv[data-test="search-title"]')
                        total_job_count = int(job_count_elem.text.replace(',', '').split()[0])
                        logger.info(f"Total jobs found for '{job_title}': {total_job_count}")
                    except Exception as count_err:
                        logger.warning(f"Could not extract total job count for {job_title}: {count_err}")
                        total_job_count = float('inf')
                    
                    # Comprehensive scrolling to load more jobs
                    max_scroll_attempts = 50  # Increased from 20
                    current_scroll_attempt = 0
                    
                    while current_scroll_attempt < max_scroll_attempts:
                        # Enhanced scrolling method
                        loaded_jobs = self.scroll_and_load_comprehensive(max_attempts=3)
                        
                        # Find all job cards after scrolling
                        job_cards = self.driver.find_elements(By.CSS_SELECTOR, 'li.JobsList_jobListItem__wjTHv[data-test="jobListing"]')
                        
                        # Track scraped jobs in this iteration
                        scraped_jobs_this_iteration = 0
                        
                        for card in job_cards:
                            try:
                                # Extract job title
                                title_elem = card.find_element(By.CSS_SELECTOR, 'a.JobCard_jobTitle__GLyJ1[data-test="job-title"]')
                                job_title_text = title_elem.text.strip()
                                
                                # Extract job URL
                                job_url = title_elem.get_attribute('href')
                                
                                # Skip duplicate jobs
                                if job_url in unique_job_urls:
                                    continue
                                unique_job_urls.add(job_url)
                                
                                # Extract company name
                                try:
                                    company_name = card.find_element(By.CSS_SELECTOR, 'span.EmployerProfile_compactEmployerName__9MGcV').text.strip()
                                except Exception:
                                    company_name = "N/A"
                                
                                # Always set location to Malaysia
                                location = "Malaysia"
                                
                                # Check for Easy Apply
                                try:
                                    easy_apply_element = card.find_element(By.CSS_SELECTOR, '.JobCard_easyApplyTag__5vlo5')
                                    easy_apply = 'Yes'
                                except:
                                    easy_apply = 'No'
                                
                                # Create job dictionary
                                job_info = {
                                    'Platform': 'Glassdoor',
                                    'Job Title': job_title_text,
                                    'Company': company_name,
                                    'Location': location,
                                    'url': job_url,
                                    'Search Keyword': job_title,
                                    'source': 'Glassdoor',
                                    'Easy Apply': easy_apply,
                                    'Salary': 'N/A'
                                }
                                
                                # Add salary extraction if available
                                try:
                                    salary = self.extract_salary(card)
                                    job_info['Salary'] = salary
                                except Exception as salary_err:
                                    logger.warning(f"Could not extract salary: {salary_err}")
                                
                                all_jobs.append(job_info)
                                scraped_jobs_this_iteration += 1
                            
                            except Exception as card_err:
                                logger.warning(f"Error processing job card: {card_err}")
                        
                        # Update scraping progress
                        logger.info(f"Scraped {len(all_jobs)} total jobs, {scraped_jobs_this_iteration} in this iteration")
                        
                        # Break if no new jobs were loaded
                        if scraped_jobs_this_iteration == 0 and loaded_jobs == 0:
                            break
                        
                        current_scroll_attempt += 1
                    
                    # Verify job count
                    self.verify_job_count(total_job_count, len(all_jobs))
                    
                    # Store job summary
                    job_summary[job_title] = {
                        'Total Jobs Found': total_job_count,
                        'Jobs Scraped': len(all_jobs)
                    }
                    
                    logger.info(f"Scraped {len(all_jobs)} jobs for '{job_title}' in Malaysia")
                
                except Exception as e:
                    logger.error(f"Error scraping jobs for '{job_title}' in Malaysia: {e}")
                    # Optional: Handle manual verification if needed
                    if self.manual_popup_handler(job_title):
                        continue
        
        # Log overall job summary
        logger.info("Job Scraping Summary:")
        for title, summary in job_summary.items():
            logger.info(f"{title}: {summary['Jobs Scraped']} of {summary['Total Jobs Found']} jobs scraped")
        
        # Export jobs to Excel
        self.export_to_excel(all_jobs)
        
        return all_jobs

    def extract_salary(self, job_card):
        try:
            # Multiple selectors for robustness
            salary_selectors = [
                '.JobCard_salaryEstimate__QpbTW',  # Main salary selector
                '.JobCard_salaryEstimateWrapper__oPNI_',  # Fallback selector
                'div[data-test="detailSalary"]'  # Alternative selector
            ]
            
            for selector in salary_selectors:
                try:
                    salary_element = job_card.find_element(By.CSS_SELECTOR, selector)
                    if salary_element:
                        salary_text = salary_element.text.strip()
                        
                        # Clean and standardize salary text
                        salary_text = salary_text.replace('(Employer est.)', '').strip()
                        
                        # Extract numeric value
                        salary_match = re.search(r'(MYR)?\s*(\d+(?:K)?)', salary_text)
                        if salary_match:
                            salary = salary_match.group(2)
                            return f"{salary_match.group(1) or 'MYR'} {salary}"
                        
                        return salary_text
                except Exception as e:
                    logging.warning(f"Error extracting salary with selector {selector}: {e}")
            
            return "Salary Not Disclosed"
        
        except Exception as e:
            logging.error(f"Comprehensive salary extraction failed: {e}")
            return "Salary Not Disclosed"

    def export_to_excel(self, jobs, filename=None):
        """
        Export scraped jobs to an Excel file
        
        :param jobs: List of job dictionaries
        :param filename: Optional custom filename, defaults to timestamped filename
        :return: Path to the exported Excel file
        """
        import pandas as pd
        from datetime import datetime
        
        # Convert jobs to DataFrame
        df = pd.DataFrame(jobs)
        
        # Create filename if not provided
        if filename is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"d:/scrape/glassdoor_jobs_{timestamp}.xlsx"
        
        # Ensure directory exists
        import os
        os.makedirs(os.path.dirname(filename), exist_ok=True)
        
        # Export to Excel
        try:
            # Reorder columns for better readability
            column_order = [
                'Job Title', 'Company', 'Location', 'Salary', 
                'url', 'Search Keyword', 'Platform', 
                'Easy Apply', 'source'
            ]
            
            # Select only columns that exist in the DataFrame
            existing_columns = [col for col in column_order if col in df.columns]
            df_ordered = df[existing_columns]
            
            # Export to Excel
            df_ordered.to_excel(filename, index=False, engine='openpyxl')
            
            logger.info(f"Jobs exported to {filename}")
            return filename
        
        except Exception as e:
            logger.error(f"Error exporting jobs to Excel: {e}")
            return None

    def save_results(self, jobs):
        """
        Save scraped job listings to Excel with multiple sheets
        
        :param jobs: List of job dictionaries
        """
        try:
            # Create timestamp for unique filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(self.output_dir, f'glassdoor_jobs_{timestamp}.xlsx')
            
            # Convert to DataFrame
            df = pd.DataFrame(jobs)
            
            # Create Excel writer
            with pd.ExcelWriter(output_file) as writer:
                # Job Listings Sheet
                df.to_excel(writer, sheet_name='Job Listings', index=False)
                
                # Summary by Job Title
                job_title_summary = df.groupby('Search Keyword').size().reset_index(name='Job Count')
                job_title_summary.to_excel(writer, sheet_name='Job Title Summary', index=False)
                
                # Summary by Company
                company_summary = df.groupby('Company').size().reset_index(name='Job Count')
                company_summary.to_excel(writer, sheet_name='Company Summary', index=False)
            
            logger.info(f"Total jobs found: {len(jobs)}")
            logger.info(f"Results saved to {output_file}")
            
            # Optional: Open the file after saving
            os.startfile(output_file)
        
        except Exception as e:
            logger.error(f"Error saving results: {e}")
            logger.error(traceback.format_exc())

    def __del__(self):
        """Close browser on object destruction"""
        try:
            if self.driver:
                logger.info("Attempting to close Glassdoor Browser...")
                self.driver.quit()
                logger.info("Glassdoor Browser closed successfully")
        except Exception as e:
            logger.error(f"Error closing browser: {e}")
            logger.error(traceback.format_exc())

def main():
    # Path to ChromeDriver
    CHROMEDRIVER_PATH = r'C:\Users\numan\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe'
    
    scraper = None
    try:
        # Initialize and run Glassdoor scraper
        scraper = GlassdoorScraper(CHROMEDRIVER_PATH)
        jobs = scraper.scrape_jobs()
        scraper.save_results(jobs)
    
    except Exception as e:
        logger.error(f"Glassdoor Scraping Error: {e}")
        logger.error(traceback.format_exc())
        print(f"An error occurred: {e}")
    
    finally:
        # Ensure driver is closed even if an exception occurs
        if scraper and scraper.driver:
            try:
                scraper.driver.quit()
            except Exception:
                pass

if __name__ == "__main__":
    main()
