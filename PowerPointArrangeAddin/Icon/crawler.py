import requests
import bs4
import time
from concurrent.futures import ThreadPoolExecutor

# ==> 2016
# url = 'https://ymrt.jp/imagemso/imagemso.cgi?wildkeys1=*&did=2016&size=24&print=500&mode=2&disp=ON&p={}'
# pages = 18

# ==> 2010
url = 'https://ymrt.jp/imagemso/imagemso.cgi?wildkeys1=*&did=2010&size=24&print=500&mode=2&disp=ON&p={}'
pages = 17

accept = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7'
cookie = 'lunalys_id=id%3D881_653d218eedbc0%26visit%3D3'
ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36'

for i in range(0, pages):
    u = url.format(i)
    resp = requests.get(u)
    doc = bs4.BeautifulSoup(resp.text, features='lxml')
    imgs = doc.select("img.ImageArea")

    with ThreadPoolExecutor(max_workers=30) as executor:
        l = len(imgs)
        for idx, img in enumerate(imgs):
            src = 'https://ymrt.jp/imagemso/' + img.attrs['src']
            alt = img.attrs['alt']

            def handle(i, idx, l, src, alt):
                print(f"[{i+1}/{pages}] [{idx + 1}/{l}] {alt}\t\t\t{src}")
                img_resp = requests.get(src, headers={
                    'Accept': accept,
                    'Cookie': cookie,
                    'User-Agent': ua
                })
                with open(f'./icon/{alt}.png', 'wb') as f:
                    f.write(img_resp.content)

            time.sleep(0.05)
            executor.submit(handle, i, idx, l, src, alt)

    time.sleep(10)
