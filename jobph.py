import requests
import pandas as pd
from bs4 import BeautifulSoup


def get_html(page_Number):
    jobPH_url = f'https://www.onlinejobs.ph/jobseekers/jobsearch/{page_Number}'
    agent = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'}
    r = requests.get(url=jobPH_url, headers=agent)
    soup = BeautifulSoup(r.content, 'html.parser')
    return soup

def scrape_information(soup):
    job_Section = soup.find_all('div', class_='jobpost-cat-box')
    for job_information in job_Section:
        
        try: 
            job_title = job_information.find('h4', class_='fs-16').text.strip()
        except Exception as e:
            job_title = " "
            print(e)
       
        try:
            salary = job_information.find('dd', class_='col').text.strip()
        except Exception as e:
            salary = " "
            print(e)
        
        try:
            company = job_information.find('p', class_='fs-13').text.strip()
        except Exception as e:
            company = " "
            print(e)
        
        try:
            date_posted = job_information.find('em').text.strip()
        except Exception as e:
            date_posted = " "
            print(e)
        
        try:
            employment_type = job_information.find('span', class_='badge').text.strip()
        except Exception as e:
            employment_type = " "
            print(e)


        job_dict = {
            'Job-Title' : job_title,
            'Salary-Offer' : salary,
            'Company' : company,
            'Date-Posted' : date_posted,
            'Type-of-Employment' : employment_type

        }
        job_List.append(job_dict)
    return
if __name__ == "__main__":
    job_List = []
    for x in range(0, 990, 30):
        job = get_html(30)
        scrape_information(job)

    df = pd.DataFrame(job_List)
    df.to_csv('Job-Hiring.csv')

    job_List_to_Excel = pd.ExcelWriter('Job-Hiring.xlsx', engine='openpyxl')
    df.to_excel(job_List_to_Excel, sheet_name='Job-Vacant', index=False)
    
    job_WorkBook = job_List_to_Excel.book
    job_WorkSheet = job_List_to_Excel.sheets['Job-Vacant']

    for column in job_WorkSheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjust_Width = (max_length + 1)
        job_WorkSheet.column_dimensions[column[0].column_letter].width = adjust_Width

    job_List_to_Excel._save()


