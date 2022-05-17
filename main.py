import sys
import time
import xlrd
from pathlib import Path
import os
import platform
import requests as r
import re as re
from shutil import copyfileobj
from PIL import Image

directory = os.getcwd()
headers = {
    'user-agent': r'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.79 Safari/537.36'
}

clean_html = re.compile(r'<.*?>|&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-f]{1,6});', re.I | re.M)
clean_url = re.compile(r'^http[s]?:\/\/(www\.)?(.*)?\/?(.)*', re.I | re.M)
clean_filename = re.compile(r"[\s\d()a-zA-Zα-ωΑ-Ω\w\d_.-]*$", re.I | re.M)
isImage = re.compile(r'jpg$|.png$|.gif$|.jpeg$', re.I | re.M)


def clean():
    if platform.system() == 'Windows':
        return os.system('cls')
    else:
        return os.system('clear')


def getXls():
    for file in os.listdir():
        if file.endswith('.xls'):
            return readXls(file)
    return print('File not found, there are no XLS files in current directory.\n'
                 'Make sure that there are .xls files in your working directory and try again..')


def readXls(fname):
    book = xlrd.open_workbook_xls(fname)
    sheet = book.sheet_by_index(0)

    rowxCounter = 1
    links = []

    try:
        while True:
            links.append(sheet.cell_value(rowx=rowxCounter, colx=3))
            rowxCounter += 1
    except IndexError:
        return getImage(links)


def getImage(urls):
    counter = 1
    imgname = []
    os_name = platform.system()
    for url in urls:
        total = len(urls)
        if re.search(isImage, str(url)):
            image = clean_url.search(str(url))
            image = image.group()

            img_url = url

            filename = clean_filename.search(str(image))
            filename = filename.group()

            for character in '!_@#${}[]|?~`/%^&*()~=:;"\'><':
                filename = filename.replace(character, '')

            filename.strip()
            if img_url.startswith('http://') or img_url.startswith('https://') and re.search(isImage, img_url):
                req = r.get(img_url, stream=True, headers=headers)
            else:
                print('\rskipped: -> ' + img_url)
                continue

            req.raw.decode_content = True

            path = f'images/{filename}' if os_name != 'Windows' else f'images\\{filename}'

            def retry():
                pass

            if not Path(path).is_file():
                pass
            else:
                print(
                    f'\rFile already exists -->   \"{filename}\"                                         ')
                counter += 1
                continue

            if req.status_code == 200:
                if not os.path.isdir('images'):
                    os.mkdir('images')
                try:
                    if os_name == 'Windows':
                        with open('images\\' + filename, 'wb') as f:
                            copyfileobj(req.raw, f)
                            req.close()
                            imgname.append(filename)
                            counter += 1
                    else:
                        with open('images/' + filename, 'wb') as f:
                            copyfileobj(req.raw, f)
                            req.close()
                            imgname.append(filename)
                            counter += 1
                except PermissionError:
                    print('Insufficient permissions.')
                except FileNotFoundError:
                    print('File not found')
                except IOError:
                    print('IOError can not read or write file.')
                except Exception:
                    print('Unknown error')

                sys.stdout.write(
                    f'\rDownloading images: [{(100 * float(counter) / float(total)):.2f}%] | {counter} / {total}')
                sys.stdout.flush()
            else:
                sys.stdout.write(                        '\rAn error has occurred, The website or your Internet connection is offline.\nretrying in 5 seconds..')
                sys.stdout.flush()
                time.sleep(5)
                retry()
    sys.stdout.flush()
    sys.stdout.write(f'\r                                                               ')
    return processImage(imgname)


def processImage(imgname):
    counter = 1
    total = len(imgname)
    os_name = platform.system()
    for image in imgname:
        if re.search(isImage, image):
            if os_name == 'Windows':
                img = Image.open('images\\' + image).convert('RGB').save('images\\' + image)
                img = Image.open('images\\' + image)
            else:
                img = Image.open('images/' + image).convert('RGB').save('images/' + image)
                img = Image.open('images/' + image)

            img_resize = img.resize((2000, 2000))

            if not os.path.isdir('processed'):
                os.mkdir('processed')

            if os_name == 'Windows':
                img_resize.save(f'processed\\{image}')
            else:
                img_resize.save(f'processed/{image}')

            sys.stdout.write(f'\rProcessing images: [{100 * float(counter) / float(total):.2f}]% | {counter} / {total}')
            sys.stdout.flush()
            counter += 1
        else:
            pass
    sys.stdout.flush()
    sys.stdout.write('\r                                                                             ')
    sys.stdout.write('\rCompleted')
    sys.exit()


if __name__ == '__main__':
    clean()
    getXls()
