import hashlib, requests, datetime
from bs4 import BeautifulSoup as bs
from tqdm.auto import tqdm

from requests.exceptions import SSLError, MissingSchema, InvalidSchema

# header section for requests package. Helps to avoid 403 Forbidden error
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}

# navigate excel file using column names 
columns = {"link": 1, 
           "format": 2, 
           "last_updated": 3, 
           "html_tag":4, 
           "hash":5, 
           "condition":6,
           "status_code": 7,
           "reason": 8}


def check_link(url:str, headers:str=headers):
    """check the status of the web link (URI)

    Args:
        url (str): link
        headers (str, optional): browser head]ers. Defaults to headers.

    Returns:
        status (str): link status code
        reason (str): status code explanation
    """
    request, status, reason = None, None, None
    try:
        request = requests.head(url, allow_redirects=False, verify=True, headers=headers)
        status = request.status_code
        reason = request.reason
        assert status == 200      
    except SSLError: 
        # nothing to worry about, just pass 
        pass
    except (MissingSchema, InvalidSchema):
        # tqdm.write(f">Invalid: {url}")
        pass
    finally:
        return status, reason
    

def get_content(url:str, div:str="", headers:str=headers)->list:
    """extract content inside <p> tag. consider specific <div> if specified.

    Args:
        url (str): url to extract content from
        div (str, optional): <div> tag to search through. Defaults to "".
        headers (str, optional): requests header section (avoids 403 error). Defaults to headers.

    Returns:
        list: content inside <p> tags
    """
    content = []
    request = requests.get(url, allow_redirects=False, headers=headers)
    html = bs(request.text, 'html.parser')
    html_divs = html.find_all("div", {"class": div})
    # loop through <div> and <p> tags
    for div in html_divs:
        paras = div.find_all('p')
        for p in paras:
            content.append(p.getText())     
               
    return content
    

def gen_hash(data:str)->str:
    """generate hash from given string data

    Args:
        data (str): data to generate hash

    Returns:
        str: hash
    """
    encoded_data = data.encode()
    hash = hashlib.sha224(encoded_data).hexdigest()
    
    return hash

def update_cell(worksheet, row, column_name, value, columns=columns)->None:
    worksheet.cell(row=row, column=columns[column_name]).value = value
    

def execute_row(worksheet, row:int):
    # extract URI
    time = datetime.datetime.now()
    try:
        url = worksheet.cell(row=row, column=columns["link"]).hyperlink.target
        status, reason = check_link(url)
        if status == None:
            raise AttributeError
    except AttributeError:
        update_cell(worksheet, row, "last_updated", time)
        update_cell(worksheet, row, "condition", "Invalid Link")
        # tqdm.write(f">Error in the URL: {url} ({status} - {reason})") 
        return False # break
    
    if status:
        # Broken Link
        update_cell(worksheet, row, "last_updated", time)
        update_cell(worksheet, row, "status_code", status)      
        update_cell(worksheet, row, "reason", reason)
        
    if status >= 400:
        update_cell(worksheet, row, "condition", "Broken Link")
    
    if status < 400:
        div = worksheet.cell(row=row, column=columns["html_tag"]).value
        # empty cell returns NoneType
        if not div:
            div=""
    
        # generate hash - 1st time
        data = get_content(url, div)
        hash1 = gen_hash(str(data))
    
        # generate hash - 2nd time
        data = get_content(url, div)
        hash2 = gen_hash(str(data))
    
        old_hash = worksheet.cell(row=row, column=columns["hash"]).value
        if (not old_hash) and (hash1 == hash2):
            update_cell(worksheet, row, "hash", hash1)
            update_cell(worksheet, row, "condition", "Just added")
            update_cell(worksheet, row, "last_updated", time)
            return True
    
        # hashes generated with in a short timeframe do not match -> web page has dynamic content
        if not hash1 == hash2:
            update_cell(worksheet, row, "hash", "ERROR")
            return False
    
        if hash1 == old_hash:
            update_cell(worksheet, row, "condition", "No Change")
            update_cell(worksheet, row, "last_updated", time)
            return True
        else:
            update_cell(worksheet, row, "condition", "Updated")
            update_cell(worksheet, row, "hash", hash1)
            update_cell(worksheet, row, "last_updated", time)

