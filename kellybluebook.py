import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import time
import re
import numpy as np
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium.common.exceptions import TimeoutException


results = []

known_makes = [
        "Acure", "Aston Martin", "Audi", "BMW", "Bentley", "Buick", "Cadillac", "Chevrolet",
        "Chrysler", "Dodge", "Ferrari", "Ford", "GMC", "HUMMER", "Honda", "Hyundai", 
        "INFINITI", "Isuzu", "Jaguar", "Jeep", "Kia", "Lamborghini", "Land Rover", "Lexus",
        "Lincoln", "Lotus", "MAZDA", "MINI", "Maserati", "Maybach", "Mercedes-Benz",
        "Mercury", "Mitsubishi", "Nissan", "Panoz", "Pontiac", "Porsche", "Rolls-Royce",
        "Saab", "Saturn", "Scion", "Subaru", "Suzuki", "Toyota", "Volkswagen", "Volvo"
]

colors_priority = [
    "colorPicker_Black",
    "colorPicker_White",
    "colorPicker_Gray",
    "colorPicker_Silver",
    "colorPicker_Red",
    "colorPicker_Blue",
    "colorPicker_Beige",
    "colorPicker_Brown",
    "colorPicker_Burgundy",
    "colorPicker_Gold",
    "colorPicker_Pink",
    "colorPicker_Green",
    "colorPicker_Orange",
    "colorPicker_Purple",  
    "colorPicker_Yellow" 
]

df = pd.read_excel("/Users/janiragayle/Desktop/Kelly Blue Book/test_sheet.xlsx")

def split_make_model(car_make_model, known_makes):
        car_make_model = str(car_make_model).strip()
        
        #if empty or 97.no answer
        if not car_make_model or "97. No Answer" in car_make_model:
                return (np.nan, np.nan) 
        
        #splits make and model
        for make in known_makes:
                if car_make_model.lower().startswith(make.lower()):
                        model = car_make_model[len(make):].strip()
                        return (make, model)
                
        return (np.nan, np.nan) #safety measure

def check_year(year):

        if pd.isna(year):
                return year
        if isinstance(year, str):
                if "97. No Answer" in year:
                        return np.nan 
        else:
                return(str(int(year)))

def calc_mileage(year):
        # We calculate miles by multiplying 5,000 * (Year of Survey Administration – Year of Car). 
        #For the 27M spreadsheet I shared, the survey year is 2017.
        survey_year = 2024
        mileage = str(5000 * (survey_year - int(year)))
        print(mileage)
        return mileage


def get_trade_in_value(make, model, year, mileage):
        
        #check if make/model or year is NaN
        if pd.isna(make) or pd.isna(model) or pd.isna(year):
                return np.nan #skips rows lacking information

        #checks if make/model or year is 97. no answer
        if "No Answer" in str(make) or "No Answer" in str(model) or "No Answer" in str(year):
                return np.nan
        
        #grabs make/model element
        try:
                # Initialize the WebDriver service
                service = Service('/Users/janiragayle/Desktop/Kelly Blue Book/chromedriver')
                driver = webdriver.Chrome(service=service)

                # Navigate to a website
                driver.get("https://www.kbb.com/whats-my-car-worth/")

                #gives time for website to load fully
                time.sleep(2)


                #grabs make/model element

                try: 
                        make_model_element = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Make/Model']"))
                        )
                        driver.execute_script("arguments[0].scrollIntoView();", make_model_element)
                        make_model_element.click()

                except TimeoutException:
                        print("make/model dial button not clicked")
                        return np.nan

                #input year
                try: 

                        wait = WebDriverWait(driver, 10)
                        select_year_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Year']")))
                        
                        year_dropdown = Select(select_year_element)

                        wait.until(EC.element_to_be_clickable((By.XPATH, f"//select[@aria-label='Year']/option[text()='{year}']")))
                        
                        year_dropdown.select_by_visible_text(year)

                except TimeoutException:
                        print("year not clicked") 
                        return np.nan               

                #input make
                try: 
                        wait = WebDriverWait(driver, 10)
                        select_make_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Make']")))
                        
                        make_dropdown = Select(select_make_element)

                        wait.until(EC.element_to_be_clickable((By.XPATH, f"//select[@aria-label='Make']/option[text()='{make}']")))
                        
                        make_dropdown.select_by_visible_text(make)

                #input model
                except TimeoutException:
                        print("make not clicked") 
                        return np.nan                  


                try: 
                        wait = WebDriverWait(driver, 10)
                        select_model_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[aria-label='Model']")))
                        
                        model_dropdown = Select(select_model_element)

                        wait.until(EC.element_to_be_clickable((By.XPATH, f"//select[@aria-label='Model']/option[text()='{model}']")))
                        
                        model_dropdown.select_by_visible_text(model)

                except TimeoutException:
                        print("model not clicked")
                        return np.nan                

                #input mileage
                try: 
                        wait = WebDriverWait(driver, 10)
                        mileage_element = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "mileage")))

                        mileage_element.send_keys(mileage)

                except TimeoutException:
                        print("unable to input mileage")
                        return np.nan

                """time.sleep(2)
                submit_element = driver.find_element(By.CSS_SELECTOR, '[data-automation="submit-button"]')
                submit_element.click()"""

                #click submit
                try: 
                        wait = WebDriverWait(driver, 10)
                        submit_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-automation="submit-button"]')))

                        submit_element.click()     
                except TimeoutException:
                        print("unable to click submit #1") 
                        return np.nan               

                #if coupe option appears
                try:
                        coupe_element = WebDriverWait(driver, 10).until(
                                 EC.presence_of_element_located((By.CSS_SELECTOR, '[data-lean-auto="category-1"]'))
                        )
                        coupe_element.click() 
                        print("coupe category 1 works")
                        
                except TimeoutException:
                        print("coupe text option timeout, trying image alt")

                        try:
                                coupe_element = WebDriverWait(driver, 10).until(
                                 EC.presence_of_element_located((By.CSS_SELECTOR, 'object[alt*="Coupe image"]')))
                                coupe_element.click()  

                        except TimeoutException:
                                print("coupe photo option timeout, script continues")
                        



                """time.sleep(2)
                lowest_priced_div = driver.find_element(By.XPATH, "//div[contains(text(), '(Lowest priced)')]")
                lowest_priced_div.click()
                """
                try: 
                        wait = WebDriverWait(driver, 10)
                        lowest_priced_div = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), '(Lowest priced)')]")))

                        lowest_priced_div.click()    
                except TimeoutException:
                        print("unable to click lowest price option") 
                        return np.nan  
                except NoSuchElementException:
                        print("(Lowest priced) option no such element, script continues")
                        return np.na              
                

                """time.sleep(2)
                submit_element = driver.find_element(By.CSS_SELECTOR, '[data-lean-auto="categoryPickerButton"]')
                submit_element.click()"""

                #
                try: 
                        wait = WebDriverWait(driver, 10)
                        submit_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-lean-auto="categoryPickerButton"]')))
                        submit_element.click()   
                except TimeoutException:
                        print("unable to click submit #2")
                        return np.nan                  
                
                

                """time.sleep(2)
                equipment_div = driver.find_element(By.XPATH, "//div[contains(text(), 'Price with standard equipment')]")
                equipment_div.click()"""

                #click price with standard equip
                try: 
                        wait = WebDriverWait(driver, 10)
                        equipment_div = wait.until(EC.element_to_be_clickable((By.ID, "pricestandard")))
                        equipment_div.click()   
                except TimeoutException:
                        print("unable to click with standard price using id, trying text'") 
                        try: 
                                wait = WebDriverWait(driver, 10)
                                equipment_div = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Price with standard equipment')]")))
                                equipment_div.click()   
                        except TimeoutException:
                                print("unable to click with standard price text. exiting") 
                                return np.nan 

                #click color of car
                for color in colors_priority:
                        try: 
                                color_element = WebDriverWait(driver, 2).until(
                                        EC.element_to_be_clickable((By.CSS_SELECTOR, f'[data-lean-auto="{color}"]'))
                                )
                                print(color)
                                color_element.click()
                                break
                        except TimeoutException:
                                print("error: could not pick color")
                                return np.nan


                try: 
                        wait = WebDriverWait(driver, 10)
                        trading_div = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Trading in my car')]")))
                        trading_div.click()   
                except TimeoutException:
                        print("unable to click trading in my car") 
                        return np.nan              


                try: 
                        wait = WebDriverWait(driver, 10)
                        submit_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-lean-auto="optionsNextButton"]')))
                        submit_element.click()   
                except TimeoutException:
                        print("unable to click submit/next button #3")
                        return np.nan   


                try: 
                        wait = WebDriverWait(driver, 10)
                        submit_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-lean-auto="next-btn"]')))
                        submit_element.click()   
                except TimeoutException:
                        print("unable to click submit/next button #4") 
                        return np.nan  


                try: 
                        wait = WebDriverWait(driver, 10)
                        trading_div = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-lean-auto="good"]')))
                        trading_div.click()   
                except TimeoutException:
                        print("unable to click good option") 
                        return np.nan


                try: 
                        wait = WebDriverWait(driver, 10)
                        submit_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-lean-auto="conditionNextButton"]')))
                        submit_element.click()   
                except TimeoutException:
                        print("unable to submit #5")
                        return np.nan


                """time.sleep(2)
                submit_element = driver.find_element(By.CSS_SELECTOR, '[data-lean-auto="quick-link"]')
                submit_element.click()"""

                try: 
                        wait = WebDriverWait(driver, 10)
                        submit_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-lean-auto="quick-link"]')))
                        submit_element.click()   
                except (WebDriverException, NoSuchElementException, TimeoutException):
                        print("unable to click quick link")
                        return np.nan
                
                wait = WebDriverWait(driver, 10)
                money_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'object[alt*="trade-in range"]')))

                alt_text = money_element.get_attribute('alt')
                        #print("Raw ALT text:", alt_text)

                pattern = r"trade-in value\s*\$[\d,]+"
                match = re.search(pattern, alt_text)

                if match:
                        # remove "trade-in value" to isolate the final dollar figure
                        trade_in_value = match.group().replace("trade-in value", "").strip()                      
                        print(f"{make}, {model}, {year} - {trade_in_value}")
                else:
                        print("No match found in the alt text.")
                        trade_in_value = np.nan

                #correct opion preselected, click next
                #select good, then click next
                #retrieve text value for trade in
                # Close the browser
                driver.quit()

                #trade_in_value = "$17,000"
                return trade_in_value
        
        except (WebDriverException, NoSuchElementException) as e:
                # Handle Selenium errors (e.g., element not found, timeout, etc.)
                print(f"Error occurred: {e}. next iteration")
   
        finally:
        # Ensure the driver is quit
                if driver:
                        driver.quit()    
        return np.nan  # Return NaN on error

for idx, row in df.iterrows():
        analytic_id = row["analytic_ID"]

        ###Vehicle 1###
        car_make_model1 = row["q9vehicle_model_1"]
        year1 = row["q9vehicle_year_1"]
        value1 = None

        #parse out make/model
        make1, model1 = split_make_model(car_make_model1, known_makes)
        year1 = check_year(year1)


        if not pd.isna(make1) and not pd.isna(model1) and not pd.isna(year1):
                mileage = calc_mileage(year1)
                value1 = get_trade_in_value(make1, model1, year1, mileage)
                print(f"{make1} - {model1} - {year1}")
        else: 
                pass

        ###Vehicle 2###
        car_make_model2 = row["q9vehicle_model_2"]
        year2 = row["q9vehicle_year_2"]
        value2 = None

        #parse out make/model
        make2, model2 = split_make_model(car_make_model2, known_makes)
        year2 = check_year(year2)
        

        if not pd.isna(make2) and not pd.isna(model2) and not pd.isna(year2):
                mileage = calc_mileage(year2)
                value2 = get_trade_in_value(make2, model2, year2, mileage)
                print(f"{make2} - {model2} - {year2}")
        else: 
                pass

        ###Vehicle 3###
        car_make_model3 = row["q9vehicle_model_3"]
        year3 = row["q9vehicle_year_3"]
        value3 = None

        #parse out make/model
        make3, model3 = split_make_model(car_make_model3, known_makes)
        year3 = check_year(year3)


        if not pd.isna(make3) and not pd.isna(model3) and not pd.isna(year3):
                mileage = calc_mileage(year3)
                value3 = get_trade_in_value(make3, model3, year3, mileage)
                print(f"{make3} - {model3} - {year3}")
        else: 
                pass

        if pd.isna(value1) and pd.isna(value2) and pd.isna(value3):
        # means we have no valid trade-in values
                continue

        results.append({
        "analytic_ID": analytic_id,
        "vehicle1_value": value1,  # might be None or $somevalue
        "vehicle2_value": value2,
        "vehicle3_value": value3
        })

if results:
        df_out = pd.DataFrame(results)
        df_out.to_excel("/Users/janiragayle/Desktop/Kelly Blue Book/final_values.xlsx", index=False)

