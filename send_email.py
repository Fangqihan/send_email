import requests
import xlrd


def get_to_str(file_path):
    data = xlrd.open_workbook(file_path)
    table = data.sheet_by_name('Sheet1')
    lst = []
    rows = table.nrows
    for i in range(0, rows):
        lst.append(str(table.row_values(i)[0]).replace('.0', '@qq.com'))
    s = ';'.join(lst)
    # print(s)
    return s


if __name__ == '__main__':
    url = "http://api.sendcloud.net/apiv2/mail/send"
    file_path = r'C:\Users\Administrator\PycharmProjects\review\子云\test.xlsx'

    to_str = get_to_str(file_path)

    print(to_str)

    API_USER = 'chengziyun_test_9FbRe3'
    API_KEY = 'BqrbMvQL9qjpiGbA'

    params = {
        "apiUser": API_USER, # 使用api_user和api_key进行验证
        "apiKey" : API_KEY,
        "to" : '105414915@qq.com;1002268753@qq.com', # 收件人地址, 用正确邮件地址替代, 多个地址用';'分隔
        "from" : "dashantj@zf06lX6xoH5MzKsp47UIMUW3FsuA0vdS.sendcloud.org", # 发信人, 用正确邮件地址替代
        "fromName" : "大山团建",
        "subject" : "武汉中小企业免费团建",
        "html": '''
    <div style="font-weight: 700">
        <p>武汉大山团建限时<span style="font-size: 30px;font-weight: 800;color: #1ed115">免费</span>团建活动开始啦，无需门槛，只要<span style="font-size: 30px;
        font-weight: 800;color: #1ed115">提前预定</span>就可以了。</p>
        <p>还有更多超低体验价徒步团建活动，具体请进入官网了解：<a href="http://www.odashan.com" target="_blank">http://www.odashan.com</a></p>
        <p>我们的服务宗旨是让小伙伴们在开心愉悦的氛围中去主动参与、互动、协作，以此来提升团队凝聚力和对企业的归属感！</p>
        
        <div>
            <p style="font-size: 10px;color: #9da4aa">若下方图片未显示，请点击上方黄色区域 &nbsp;显示图片</span></p>
            <p><img src="http://oyhijg3iv.bkt.clouddn.com/11.jpg"></p>
            <p><img src="http://oyhijg3iv.bkt.clouddn.com/3.jpg"></p>
            <p><img src="http://oyhijg3iv.bkt.clouddn.com/5.jpg"></p>
        </div>
        
        <div>
            <img src="http://oyhijg3iv.bkt.clouddn.com/6.jpg">
            <p>微信公众号：大山探游团建 </p>
            <p>联系人：程先生&nbsp;&nbsp;&nbsp;联系电话：18071118383</p>
        </div>
    </div>
    '''
    }

    r = requests.post(url, data=params)
    print(r.text)
