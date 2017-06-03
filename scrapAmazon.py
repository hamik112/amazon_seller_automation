
from bs4 import BeautifulSoup
import requests
import time
import json
import datetime

a=datetime.date.today()
out = open("out-%s-%s-%s.csv"%(a.day,a.month,a.year),"w")
with open("urls.txt") as file:
    for asin in file.readlines():
        headers = {'user-agent': 'Mozilla/5.0'}
        with requests.Session() as s:
            tries = 0
            c = 0
            while c <=1 and tries <10:
                req = s.get("http://www.amazon.in/gp/product/"+asin, headers=headers)
                soup = BeautifulSoup(req.content,"html.parser")
                bsrs = soup.findAll("li",{"class":"zg_hrsr_item"})
                bsrs = [k.text.replace("\n"," ").replace("\xa0"," ").strip() for k in bsrs]
                d = {e['name']: e.get('value', '') for e in soup.find("form",{"id":"addToCart"}).find_all('input', {'name': True})}
                
                r = s.post("http://www.amazon.in/gp/product/handle-buy-box/ref=dp_start-bbf_1_glance",data=d, headers=headers)
                
                
                time.sleep(1)
                
                req = s.get("https://www.amazon.in/gp/cart/view.html/ref=nav_cart", headers=headers)
                soup = BeautifulSoup(req.content,"html.parser")
                dd = {e['name']: e.get('value', '') for e in soup.find("form",{"id":"activeCartViewForm"}).find_all('input', {'name': True})}
                id = [k for k in dd.keys() if "quantity." in k][-1].split(".")[-1]
                dd["quantity."+id]=999
                dd["quantityBox"]=999
                r = s.post("https://www.amazon.in/gp/cart/ajax-update.html/ref=ox_sc_update_quantity_1",data=dd, headers=headers)
                req = r.content
                
                results= s.get("https://www.amazon.in/gp/navigation/ajax/dynamic-menu.html", headers=headers).content
                cc = int(json.loads(results.strip().decode())["count"])
                c = cc if cc>c else c
                tries+=1
        out.write("%s,%s,%s, - ,%s\n"%(asin,"%s-%s-%s"%(a.day,a.month,a.year),c,",".join(bsrs)))
        print("%s,%s,%s, - ,%s\n"%(asin,"%s-%s-%s"%(a.day,a.month,a.year),c,",".join(bsrs)))
out.close()