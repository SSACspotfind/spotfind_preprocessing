import requests
import urllib.request
import os
from tourAPI_photogallery.resource.keys import ServiceKey


def geturl(pageno):
    dataJson = "http://api.visitkorea.or.kr/openapi/service/rest/PhotoGalleryService/galleryList?ServiceKey=" + ServiceKey + "&numOfRows=10&pageNo="+str(pageno)+"&MobileOS=ETC&MobileApp=TestApp&_type=json"
    response = requests.get(dataJson)
    data = response.json()
    count = data['response']['body']['numOfRows']
    # 전체 이미지 수  4329
    print(data['response']['body']['totalCount'])
    
    # 이미지 다운로드
    maketourImg(count , data)

def maketourImg(count , data):
    # 이름 , 장소 , 이미지
    for i in range(count):
        print(data['response']['body']['items']['item'][i]['galTitle'])
        image_title = data['response']['body']['items']['item'][i]['galTitle']
        print(data['response']['body']['items']['item'][i]['galPhotographyLocation'])
        print(data['response']['body']['items']['item'][i]['galWebImageUrl'])
        image_url = data['response']['body']['items']['item'][i]['galWebImageUrl']
        # 폴더 생성하기
        os.makedirs('./tourapi_img/' + image_title + '/', exist_ok=True)
        # image 파일 저장
        urllib.request.urlretrieve(image_url, './tourapi_img/' + image_title + '/' + image_title + '_main.jpg')

if __name__ == '__main__':
    geturl(1)
    # 전체 다운로드
    # for i in range(1,433) :
    #     geturl(i)