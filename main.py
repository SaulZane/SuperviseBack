# 创建一个FastAPI服务器
import logging
from datetime import datetime
from typing import Optional

import cx_Oracle
import uvicorn
from docxtpl import DocxTemplate
from fastapi import FastAPI
from fastapi import Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from starlette.staticfiles import StaticFiles

app = FastAPI()

# 设置跨域
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class ReportData(BaseModel):
    """
    报表数据
    """
    dailya: int
    dailyp: int
    yeara: int
    dailydriver: int
    yeardriver: int
    reporttime: datetime
    bz: Optional[str]


def get_db_conn():
    # 连接数据库
    # conn = sqlite3.connect('Daily_Report.db')
    # 上线时真正的数据库
    # todo here
    dsn = cx_Oracle.makedsn("10.88.92.25", 1521, service_name="orcl")
    conn = cx_Oracle.connect(user="vehweb", password="#weB%0413", dsn=dsn)
    # 创建游标
    cursor = conn.cursor()
    # 返回游标
    return cursor


logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', filename='server.log')


# 创建一个路由，路径为"/"
@app.get("/")
# 创建一个函数，函数名为root
def root(request: Request):
    # 返回一个字符串
    host = request.client.host
    port = request.client.port
    logging.debug(f"{host}:{port}访问了根路径")
    return {"message": "Hello World"}


# 创建一个路由，路径为"/getDataList"，接收一个参数，参数类型为Date，返回查询出的字典数据
@app.get("/getDataList")
def getdatalist(date: str, request: Request):
    """
    查询某一天的报表最新数据
    :param date: 要查询的日期
    :return: 将数据库中的内容字段显示出来，对于没有的字段返回无数据
    """
    # 删除static/工作日报.docx这个文件
    import os
    if os.path.exists("./static/工作日报.docx"):
        os.remove("./static/工作日报.docx")
    host = request.client.host
    port = request.client.port
    # 把传入的字符串yyyy/MM/dd转换成日期格式,并且改为yyyy-MM-dd
    date: date = datetime.strptime(date, '%Y/%m/%d').date()
    sentence = f"select * from  ZS_DAILY_REPORT where reporttime=(select  max(reporttime) from ZS_DAILY_REPORT where  reporttime>=date '{date}'  and   reporttime<=date '{date}'+1)"
    logging.debug(sentence)
    # 查询数据库
    results = get_db_conn().execute(sentence).fetchone()
    if results is None:
        logging.debug(f"{host}:{port}没有记录{date}的报表数据")
        return {f"没有记录{date}的报表数据"}
    else:
        # 将查询出的数据赋值给变量,解包
        dailya, dailyp, yeara, dailydriver, yeardriver, reporttime, bz = results
        # 将数据封装成字典,适配pydantic
        export_data = ReportData(dailya=dailya, dailyp=dailyp, yeara=yeara, dailydriver=dailydriver,
                                yeardriver=yeardriver, reporttime=reporttime, bz=bz)
        expordocx(export_data)
        logging.debug(f"{host}:{port}导出{export_data}报表成功")
        return {'dailya': dailya, 'dailyp': dailyp, 'yeara': yeara, 'dailydriver': dailydriver,
                'yeardriver': yeardriver,
                'reporttime': reporttime, 'bz': bz}


def expordocx(reportdata: ReportData):
    """
    导出报表
    """

    doc = DocxTemplate("./static/template/工作日报模板.docx")
    context = {'dailya': reportdata.dailya, 'dailyp': reportdata.dailyp, 'yeara': reportdata.yeara,
               'dailydriver': reportdata.dailydriver,
               'yeardriver': reportdata.yeardriver, 'reporttime': reportdata.reporttime, 'bz': reportdata.bz,
               'month': reportdata.reporttime.month, 'day': reportdata.reporttime.day,
               'exporttime': datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
    doc.render(context)
    doc.save("./static/工作日报.docx")


# 访问路径为/static/工作日报.docx
app.mount("/static", StaticFiles(directory="static"), name="static")  #

# 运行服务器
if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)  # reload=True表示修改代码后自动重启服务器
    logging.debug("服务器启动成功")
