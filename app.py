import io
from datetime import date
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="NQ Mint 收益管理仪表盘", page_icon="📊", layout="wide")
st.markdown("""<style>
.block-container{padding-top:1.2rem}.kpi{border:1px solid #dfe7ef;border-radius:14px;padding:14px;background:#fff}
[data-testid="stMetricValue"]{font-size:1.55rem}.stTabs [data-baseweb="tab"]{font-weight:700}
</style>""",unsafe_allow_html=True)
st.title("NQ Mint 收益管理仪表盘")
st.caption("单文件更新｜Actual / OTB｜FY26 vs FY25｜Revenue · RN · ADR · Occupancy · RevPAR")

DATE='Stay Date'; RN='Current Room Nights'; REV='Current Room Revenue'; OCC='Occupancy %'; REVPAR='Current Room RevPAR'
DIMS=['Property Name','Market Category','Market Segment','Rate Program','Channel Aggregate','Channel Details','Channel Partner','Loyalty Level','Travel Agency Name','Source Country/Region','Weekday Vs Weekend']
REQ=[DATE,RN,REV]

def load(raw):
    x=pd.read_excel(io.BytesIO(raw),sheet_name='page',engine='openpyxl')
    x.columns=x.columns.astype(str).str.strip()
    miss=[c for c in REQ if c not in x]
    if miss: raise ValueError('缺少必要字段：'+', '.join(miss))
    x[DATE]=pd.to_datetime(x[DATE],errors='coerce')
    x=x[x[DATE].notna()].copy()
    for c in [RN,REV,OCC,REVPAR,'Total Current RevPAR','Current Package Revenue','Addon Revenue','Total Fees','Total Tax']:
        if c in x: x[c]=pd.to_numeric(x[c],errors='coerce').fillna(0)
    x['FY']=x[DATE].dt.year.astype(int)
    x['Month']=x[DATE].dt.strftime('%m月')
    x['Month No']=x[DATE].dt.month
    return x

def val(x,m):
    if x.empty:return np.nan
    if m=='Room Revenue': return x[REV].sum()
    if m=='Room Nights': return x[RN].sum()
    if m=='ADR': return x[REV].sum()/x[RN].sum() if x[RN].sum() else np.nan
    if m=='Occupancy %': return x[OCC].mean() if OCC in x else np.nan
    return x[REVPAR].mean() if REVPAR in x else np.nan

def grouped(x,by,m):
    if x.empty:return pd.DataFrame(columns=[by,'值'])
    if m=='ADR':
        g=x.groupby(by,as_index=False).agg(_r=(REV,'sum'),_n=(RN,'sum'));g['值']=g['_r']/g['_n'].replace(0,np.nan)
    elif m=='Room Revenue': g=x.groupby(by,as_index=False)[REV].sum().rename(columns={REV:'值'})
    elif m=='Room Nights': g=x.groupby(by,as_index=False)[RN].sum().rename(columns={RN:'值'})
    elif m=='Occupancy %': g=x.groupby(by,as_index=False)[OCC].mean().rename(columns={OCC:'值'})
    else:g=x.groupby(by,as_index=False)[REVPAR].mean().rename(columns={REVPAR:'值'})
    return g[[by,'值']]

def fmt(v,m):
    if pd.isna(v):return '—'
    if m=='Occupancy %':return f'{v:.1%}'
    if m=='Room Nights':return f'{v:,.0f}'
    return f'¥{v:,.0f}' if m in ['Room Revenue','ADR','RevPAR'] else f'{v:,.1f}'

up=st.sidebar.file_uploader('上传最新 NQ mint 文档',type=['xlsx'])
if up is None:
    st.info('请在左侧上传 NQ mint Excel。后续每次只需上传一份同结构文档即可更新。')
    st.stop()
try: df=load(up.getvalue())
except Exception as e: st.error(f'文件读取失败：{e}'); st.stop()

st.sidebar.header('分析口径')
min_d,max_d=df[DATE].min().date(),df[DATE].max().date()
asof=st.sidebar.date_input('Actual / OTB 分界日',value=min(date.today(),max_d),help='Stay Date ≤ 分界日为 Actual；之后为 OTB。')
df['Status']=np.where(df[DATE].dt.date<=asof,'Actual','OTB')
status=st.sidebar.multiselect('状态',['Actual','OTB'],default=['Actual','OTB'])
years=sorted(df['FY'].unique(),reverse=True)
fys=st.sidebar.multiselect('财年',years,default=[y for y in [2026,2025] if y in years] or years)
drange=st.sidebar.date_input('入住日期范围',value=(min_d,max_d),min_value=min_d,max_value=max_d)
metric=st.sidebar.selectbox('核心指标',['Room Revenue','Room Nights','ADR','Occupancy %','RevPAR'])
avail=[c for c in DIMS if c in df]
breakdown=st.sidebar.selectbox('拆分维度',avail,index=avail.index('Market Segment') if 'Market Segment' in avail else 0)
filter_dims=st.sidebar.multiselect('增加筛选维度',avail,default=[c for c in ['Market Segment','Channel Aggregate','Loyalty Level'] if c in avail])
sel={}
for c in filter_dims:
    opts=sorted(df[c].fillna('(空白)').astype(str).unique())
    sel[c]=st.sidebar.multiselect(c,opts)

x=df[df['Status'].isin(status)&df['FY'].isin(fys)].copy()
if len(drange)==2:x=x[x[DATE].between(pd.Timestamp(drange[0]),pd.Timestamp(drange[1]))]
for c,v in sel.items():
    if v:x=x[x[c].fillna('(空白)').astype(str).isin(v)]

actual=x[x.Status=='Actual'];otb=x[x.Status=='OTB']
ks=st.columns(5)
for col,m in zip(ks,['Room Revenue','Room Nights','ADR','Occupancy %','RevPAR']):
    av,ov=val(actual,m),val(otb,m)
    col.metric(m,fmt(av,m),f'OTB {fmt(ov,m)}')

if not ({2025,2026}<=set(df.FY.unique())):
    missing=sorted({2025,2026}-set(df.FY.unique()))
    st.warning('当前上传文件未同时包含 FY25 与 FY26；缺少 '+', '.join(map(str,missing))+'。FY对比页将在单一上传文档同时包含两个财年后自动显示完整同比。')

t1,t2,t3,t4=st.tabs(['管理层总览','FY26 vs FY25','渠道与市场','明细与导出'])
with t1:
    c1,c2=st.columns([1.55,1])
    with c1:
        daily=[]
        for s in ['Actual','OTB']:
            g=grouped(x[x.Status==s],DATE,metric);g['状态']=s;daily.append(g)
        trend=pd.concat(daily,ignore_index=True)
        fig=px.line(trend,x=DATE,y='值',color='状态',color_discrete_map={'Actual':'#0B6E75','OTB':'#F59E0B'},title=f'{metric} 日趋势',template='plotly_white')
        st.plotly_chart(fig,use_container_width=True)
    with c2:
        b=grouped(x.assign(**{breakdown:x[breakdown].fillna('(空白)').astype(str)}),breakdown,metric).nlargest(15,'值').sort_values('值')
        fig=px.bar(b,x='值',y=breakdown,orientation='h',title=f'{breakdown} Top 15',color_discrete_sequence=['#1F4E78'],template='plotly_white')
        st.plotly_chart(fig,use_container_width=True)
    mix=x.groupby(['FY','Status'],as_index=False).agg(Revenue=(REV,'sum'),RN=(RN,'sum'))
    fig=px.bar(mix,x='FY',y='Revenue',color='Status',barmode='stack',title='Actual / OTB 收入结构',color_discrete_map={'Actual':'#0B6E75','OTB':'#F59E0B'},template='plotly_white')
    st.plotly_chart(fig,use_container_width=True)
with t2:
    rows=[]
    for m in ['Room Revenue','Room Nights','ADR','Occupancy %','RevPAR']:
        a=val(x[x.FY==2026],m);b=val(x[x.FY==2025],m);var=a-b if pd.notna(a) and pd.notna(b) else np.nan
        pct=var/abs(b) if pd.notna(var) and b not in [0,np.nan] else np.nan
        rows.append([m,a,b,var,pct])
    comp=pd.DataFrame(rows,columns=['指标','FY26','FY25','差异','同比%'])
    st.dataframe(comp,use_container_width=True,hide_index=True,column_config={'同比%':st.column_config.NumberColumn(format='%.1f%%')})
    gy=[]
    for y in [2025,2026]:
        g=grouped(x[x.FY==y],'Month No',metric);g['FY']=str(y);gy.append(g)
    fig=px.line(pd.concat(gy),x='Month No',y='值',color='FY',markers=True,title=f'FY26 vs FY25｜{metric} 月度趋势',template='plotly_white',color_discrete_map={'2026':'#0B6E75','2025':'#A5A5A5'})
    fig.update_xaxes(dtick=1,title='月份');st.plotly_chart(fig,use_container_width=True)
with t3:
    c1,c2=st.columns(2)
    for col,dim in zip([c1,c2],[d for d in ['Channel Aggregate','Market Segment'] if d in x][:2]):
        g=grouped(x.assign(**{dim:x[dim].fillna('(空白)').astype(str)}),dim,metric).nlargest(15,'值')
        col.plotly_chart(px.bar(g,x=dim,y='值',title=f'{dim}｜{metric}',template='plotly_white',color_discrete_sequence=['#0B6E75']),use_container_width=True)
    if 'Loyalty Level' in x:
        g=grouped(x.assign(**{'Loyalty Level':x['Loyalty Level'].fillna('(空白)').astype(str)}),'Loyalty Level',metric).nlargest(15,'值')
        st.plotly_chart(px.bar(g,x='Loyalty Level',y='值',title=f'Loyalty Level｜{metric}',template='plotly_white',color_discrete_sequence=['#F59E0B']),use_container_width=True)
with t4:
    st.dataframe(x,use_container_width=True,height=450)
    st.download_button('下载筛选后数据',x.to_csv(index=False).encode('utf-8-sig'),'NQ_mint_filtered.csv','text/csv')
    st.caption(f'当前筛选：{len(x):,} 行｜源文件：{len(df):,} 行｜口径：Stay Date ≤ {asof} 为 Actual，其后为 OTB。')
