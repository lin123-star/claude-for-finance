# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.utils import get_column_letter
wb=openpyxl.Workbook()
HF=PatternFill("solid",fgColor="1F4E79");OF=PatternFill("solid",fgColor="BDD7EE")
IF=PatternFill("solid",fgColor="F2F2F2");BF=Font(name="Arial",size=10,color="0000FF")
BK=Font(name="Arial",size=10);BD=Font(name="Arial",size=10,bold=True)
TF=Font(name="Arial",size=14,bold=True);HFR=Font(name="Arial",size=10,bold=True,color="FFFFFF")
TH=Border(left=Side('thin'),right=Side('thin'),top=Side('thin'),bottom=Side('thin'))
CA=Alignment(horizontal='center',vertical='center',wrap_text=True)
def sw(ws,r,c,v,fn=BK,fl=None,nf=None):
    cl=ws.cell(row=r,column=c,value=v);cl.font=fn
    if fl:cl.fill=fl
    if nf:cl.number_format=nf;cl.border=TH;cl.alignment=CA
def hr(ws,r,vs):
    for i,v in enumerate(vs):sw(ws,r,i+1,v,HFR,HF)
RH=[2714,2960,2543,2416,2762];CH=[1700,1870,1598,1506,1757];NH=[280,247,140,78,61.13]
GH=[RH[i]-CH[i] for i in range(5)]
RG=[0.25,0.22,0.18,0.15,0.12];GM=[0.37,0.38,0.39,0.40,0.40];EM=[0.035,0.05,0.065,0.08,0.09];DR=0.04
CX=[0.08,0.075,0.07,0.065,0.06];NW=[0.05,0.05,0.045,0.04,0.04];TR=0.15
S_=[175,189,215,221,258];A_=[185,200,208,215,244];R_=[206,228,235,242,269];F_=[-5,-8,3,-9.56,7.7]
YR=["2021A","2022A","2023A","2024A","2025A","2026E","2027E","2028E","2029E","2030E"]
RP=[]
for i in range(5):
    p=RH[4] if i==0 else RP[i-1];RP.append(round(p*(1+RG[i]),1))
def ni(i):
    if i<5:return NH[i]
    j=i-5;return round(RP[j]*EM[j]*(1-TR),1)
def rv(i):
    if i<5:return RH[i];return RP[i-5]
ws=wb.active;ws.title="IS";ws.merge_cells("A1:K1")
sw(ws,1,1,"MoonsElectric IS(RMBm)",TF);hr(ws,3,["Item"]+YR)
sw(ws,5,1,"Revenue",BD)
for i in range(5):sw(ws,5,i+2,RH[i],BF,IF,"#,##0")
for i in range(5):sw(ws,5,i+7,RP[i],nf="#,##0")
sw(ws,6,1,"  Rev%")
for i in range(10):
    if i==0:sw(ws,6,i+2,"-")
    elif i<5:sw(ws,6,i+2,round(RH[i]/RH[i-1]-1,4),nf="0.0%")
    else:sw(ws,6,i+2,RG[i-5],BF,IF,"0.0%")
sw(ws,8,1,"COGS",BD)
for i in range(5):sw(ws,8,i+2,CH[i],BF,IF,"#,##0")
for i in range(5):sw(ws,8,i+7,round(RP[i]*(1-GM[i]),1),nf="#,##0")
sw(ws,9,1,"GP",BD)
for i in range(5):sw(ws,9,i+2,GH[i],nf="#,##0")
for i in range(5):sw(ws,9,i+7,round(RP[i]*GM[i],1),nf="#,##0")
sw(ws,10,1,"  GM%")
for i in range(5):sw(ws,10,i+2,round(GH[i]/RH[i],4) if RH[i] else 0,nf="0.0%")
for i in range(5):sw(ws,10,i+7,GM[i],BF,IF,"0.0%")
sw(ws,12,1,"SGA",BD)
for i in range(5):sw(ws,12,i+2,S_[i],BF,IF,"#,##0")
for i in range(5):sw(ws,12,i+7,round(RP[i]*0.095,1),nf="#,##0")
sw(ws,13,1,"Admin",BD)
for i in range(5):sw(ws,13,i+2,A_[i],BF,IF,"#,##0")
for i in range(5):sw(ws,13,i+7,round(RP[i]*0.090,1),nf="#,##0")
sw(ws,14,1,"R&D",BD)
for i in range(5):sw(ws,14,i+2,R_[i],BF,IF,"#,##0")
for i in range(5):sw(ws,14,i+7,round(RP[i]*0.097,1),nf="#,##0")
sw(ws,15,1,"Finance",BD)
for i in range(5):sw(ws,15,i+2,F_[i],BF,IF,"#,##0")
for i in range(5):sw(ws,15,i+7,0,nf="#,##0")
sw(ws,16,1,"TotalOpEx",BD)
for i in range(5):sw(ws,16,i+2,S_[i]+A_[i]+R_[i]+F_[i],nf="#,##0")
for i in range(5):sw(ws,16,i+7,round(RP[i]*0.282,1),nf="#,##0")
sw(ws,18,1,"EBIT",BD)
for i in range(5):e=GH[i]-(S_[i]+A_[i]+R_[i]+F_[i]);sw(ws,18,i+2,round(e,1),nf="#,##0")
for i in range(5):sw(ws,18,i+7,round(RP[i]*EM[i],1),nf="#,##0")
sw(ws,19,1,"  EBIT%")
for i in range(5):e=GH[i]-(S_[i]+A_[i]+R_[i]+F_[i]);sw(ws,19,i+2,round(e/RH[i],4) if RH[i] else 0,nf="0.0%")
for i in range(5):sw(ws,19,i+7,EM[i],BF,IF,"0.0%")
sw(ws,21,1,"Tax15%")
for i in range(5):e=GH[i]-(S_[i]+A_[i]+R_[i]+F_[i]);sw(ws,21,i+2,round(max(e*TR,0),1),nf="#,##0")
for i in range(5):sw(ws,21,i+7,round(max(RP[i]*EM[i]*TR,0),1),nf="#,##0")
sw(ws,23,1,"NetProfit",BD)
for i in range(5):sw(ws,23,i+2,NH[i],BF,IF,"#,##0")
for i in range(5):sw(ws,23,i+7,round(RP[i]*EM[i]*(1-TR),1),nf="#,##0")
sw(ws,24,1,"  Net%")
for i in range(5):sw(ws,24,i+2,round(NH[i]/RH[i],4) if RH[i] else 0,nf="0.0%")
for i in range(5):sw(ws,24,i+7,round(EM[i]*(1-TR),4),nf="0.0%")
sw(ws,26,1,"EPS")
for i in range(5):sw(ws,26,i+2,round(NH[i]/419,3),nf="0.000")
for i in range(5):sw(ws,26,i+7,round(RP[i]*EM[i]*(1-TR)/419,3),nf="0.000")
for c in range(1,12):ws.column_dimensions[get_column_letter(c)].width=14
ws=wb.create_sheet("CF");ws.merge_cells("A1:K1")
sw(ws,1,1,"MoonsElectric CF(RMBm)",TF);hr(ws,3,["Item"]+YR)
sw(ws,5,1,"OpCF",BD);sw(ws,6,1,"  NI")
for i in range(10):sw(ws,6,i+2,ni(i),nf="#,##0")
sw(ws,7,1,"  DA")
for i in range(10):sw(ws,7,i+2,round(rv(i)*DR,1),nf="#,##0")
sw(ws,8,1,"  WCchg")
for i in range(10):
    if i<5:sw(ws,8,i+2,"-")
    else:j=i-5;d=rv(i)-rv(i-1);sw(ws,8,i+2,round(-d*NW[j],1),nf="#,##0")
sw(ws,9,1,"  OpCF")
for i in range(10):
    nv=ni(i);da=round(rv(i)*DR,1)
    if i<5:sw(ws,9,i+2,round(nv+da,1),nf="#,##0")
    else:j=i-5;d=rv(i)-rv(i-1);wc=round(-d*NW[j],1);sw(ws,9,i+2,round(nv+da+wc,1),fill=OF,nf="#,##0")
sw(ws,11,1,"InvCF",BD);sw(ws,12,1,"  Capex")
for i in range(5):sw(ws,12,i+2,"-")
for i in range(5):sw(ws,12,i+7,round(-rv(i+5)*CX[i],1),nf="#,##0")
sw(ws,13,1,"  InvCF")
for i in range(10):v=ws.cell(row=12,column=i+2).value;sw(ws,13,i+2,v if v else 0,fill=OF,nf="#,##0")
sw(ws,15,1,"FinCF",BD);sw(ws,16,1,"  Debt")
for i in range(10):sw(ws,16,i+2,0,nf="#,##0")
sw(ws,17,1,"  Div")
for i in range(5):sw(ws,17,i+2,round(-NH[i]*0.05,1),nf="#,##0")
for i in range(5):sw(ws,17,i+7,round(-ni(5+i)*0.10,1),nf="#,##0")
sw(ws,18,1,"  FinCF")
for i in range(10):v=(ws.cell(row=16,column=i+2).value or 0)+(ws.cell(row=17,column=i+2).value or 0);sw(ws,18,i+2,round(v,1),fill=OF,nf="#,##0")
sw(ws,20,1,"CashChg")
for i in range(10):v=(ws.cell(row=9,column=i+2).value or 0)+(ws.cell(row=13,column=i+2).value or 0)+(ws.cell(row=18,column=i+2).value or 0);sw(ws,20,i+2,round(v,1),nf="#,##0")
sw(ws,22,1,"UFCF",BD);sw(ws,23,1,"  NOPAT");sw(ws,24,1,"  +DA");sw(ws,25,1,"  -Capex");sw(ws,26,1,"  -WCchg");sw(ws,27,1,"  =UFCF")
for i in range(5):
    for r in[23,24,25,26,27]:sw(ws,r,i+2,"-")
for i in range(5):
    j=i;r0=rv(i+5);n=round(r0*EM[j]*(1-TR),1);d=round(r0*DR,1);c2=round(-r0*CX[j],1);dd=rv(i+5)-rv(i+4);w=round(-dd*NW[j],1);u=round(n+d+c2+w,1)
    sw(ws,23,i+7,n,nf="#,##0");sw(ws,24,i+7,d,nf="#,##0");sw(ws,25,i+7,c2,nf="#,##0");sw(ws,26,i+7,w,nf="#,##0");sw(ws,27,i+7,u,fill=OF,nf="#,##0")
for c in range(1,12):ws.column_dimensions[get_column_letter(c)].width=14
bs=wb.create_sheet("BS");bs.merge_cells("A1:K1")
sw(bs,1,1,"MoonsElectric BS(RMBm)",TF);hr(bs,3,["Item"]+YR)
BC=[850,920,750,800,871];BA=[700,750,680,720,760];BI=[480,520,500,530,595]
Bf=[600,650,720,780,879];BG=[500,500,500,520,515];Bs=[300,350,380,400,420]
BP=[350,380,370,390,420];Bl=[80,60,80,100,112];oca=200;onca=300;ocl=200
for ri,lb in[(5,"Cash"),(6,"AR"),(7,"Inv"),(8,"OtherCA"),(9,"TotCA"),(10,"FA"),
             (11,"GW"),(12,"OtherNCA"),(13,"TotNCA"),(14,"TotA"),(16,"STD"),(17,"AP"),
             (18,"OtherCL"),(19,"LTD"),(20,"TotL"),(21,"Equity"),(22,"L+E"),
             (24,"Ratios"),(25,"D/A"),(26,"CR"),(27,"BVPS")]:
    sw(bs,ri,1,lb,BD if ri in[5,9,13,14,20,21,22] else BK)
for i in range(5):
    sw(bs,5,i+2,BC[i],BF,IF,"#,##0");sw(bs,6,i+2,BA[i],BF,IF,"#,##0")
    sw(bs,7,i+2,BI[i],BF,IF,"#,##0");sw(bs,8,i+2,oca,BF,IF,"#,##0")
    ca=BC[i]+BA[i]+BI[i]+oca;sw(bs,9,i+2,ca,nf="#,##0")
    sw(bs,10,i+2,Bf[i],BF,IF,"#,##0");sw(bs,11,i+2,BG[i],BF,IF,"#,##0")
    sw(bs,12,i+2,onca,BF,IF,"#,##0");nca=Bf[i]+BG[i]+onca;sw(bs,13,i+2,nca,nf="#,##0")
    ta=ca+nca;sw(bs,14,i+2,ta,fill=OF,nf="#,##0")
    sw(bs,16,i+2,Bs[i],BF,IF,"#,##0");sw(bs,17,i+2,BP[i],BF,IF,"#,##0")
    sw(bs,18,i+2,ocl,BF,IF,"#,##0");sw(bs,19,i+2,Bl[i],BF,IF,"#,##0")
    tl=Bs[i]+BP[i]+ocl+Bl[i];sw(bs,20,i+2,tl,nf="#,##0");eq=ta-tl;sw(bs,21,i+2,eq,nf="#,##0");sw(bs,22,i+2,ta,fill=OF,nf="#,##0")
for i in range(5):
    cl=i+7;r0=RP[i];cdc=0
    for k in range(i+1):
        r=RP[k];n=round(r*EM[k]*(1-TR),1);d=round(r*DR,1);c2=round(-r*CX[k],1)
        pv=RH[4] if k==0 else RP[k-1];w=round(-(r-pv)*NW[k],1);dv=round(-ni(5+k)*0.10,1)
        cdc+=round(n+d+c2+w+dv,1)
    nc=round(BC[4]+cdc,1);sw(bs,5,cl,nc,nf="#,##0")
    sw(bs,6,cl,round(r0*0.28,1),nf="#,##0");sw(bs,7,cl,round(r0*0.22,1),nf="#,##0");sw(bs,8,cl,220,nf="#,##0")
    ca_p=nc+round(r0*0.28,1)+round(r0*0.22,1)+220;sw(bs,9,cl,round(ca_p,1),nf="#,##0")
    sw(bs,10,cl,round(Bf[4]+80*(i+1),1),nf="#,##0");sw(bs,11,cl,500,nf="#,##0");sw(bs,12,cl,320,nf="#,##0")
    nca_p=round(Bf[4]+80*(i+1)+500+320,1);sw(bs,13,cl,nca_p,nf="#,##0");ta_p=round(ca_p+nca_p,1);sw(bs,14,cl,ta_p,fill=OF,nf="#,##0")
    sw(bs,16,cl,420,nf="#,##0");sw(bs,17,cl,round(r0*0.15,1),nf="#,##0");sw(bs,18,cl,220,nf="#,##0");sw(bs,19,cl,112,nf="#,##0")
    tl_p=round(420+r0*0.15+220+112,1);sw(bs,20,cl,tl_p,nf="#,##0")
    eq_p=round(ta_p-tl_p,1);sw(bs,21,cl,eq_p,nf="#,##0");sw(bs,22,cl,ta_p,fill=OF,nf="#,##0")
for i in range(10):
    ca_v=sum([bs.cell(row=5,column=i+2).value or 0,bs.cell(row=6,column=i+2).value or 0,bs.cell(row=7,column=i+2).value or 0,bs.cell(row=8,column=i+2).value or 0])
    ta_v=(bs.cell(row=9,column=i+2).value or 0)+(bs.cell(row=13,column=i+2).value or 0)
    tl_v=bs.cell(row=20,column=i+2).value or 0
    cl_v=(bs.cell(row=16,column=i+2).value or 0)+(bs.cell(row=17,column=i+2).value or 0)+(bs.cell(row=18,column=i+2).value or 0)
    eq_v=bs.cell(row=21,column=i+2).value or 0
    sw(bs,25,i+2,round(tl_v/ta_v,4) if ta_v else 0,nf="0.0%")
    sw(bs,26,i+2,round(ca_v/cl_v,2) if cl_v else 0,nf="0.00")
    sw(bs,27,i+2,round(eq_v/419,2),nf="0.00")
for c in range(1,12):bs.column_dimensions[get_column_letter(c)].width=14
dc=wb.create_sheet("DCF");dc.merge_cells("A1:H1")
sw(dc,1,1,"MoonsElectric DCF",TF)
WA=0.0539;TG=0.030;SH=419;NC=339;CP=63.48
sw(dc,4,1,"Inputs",HFR,HF);sw(dc,5,1,"WACC",BF,IF);sw(dc,5,2,WA,BF,IF,"0.00%")
sw(dc,6,1,"TermG",BF,IF);sw(dc,6,2,TG,BF,IF,"0.0%")
sw(dc,7,1,"NetCash",BF,IF);sw(dc,7,2,NC,BF,IF,"#,##0")
sw(dc,8,1,"SharesM",BF,IF);sw(dc,8,2,SH,BF,IF,"#,##0")
sw(dc,9,1,"Price",BF,IF);sw(dc,9,2,CP,BF,IF,"0.00")
sw(dc,11,1,"FCF(Base)",HFR,HF);hr(dc,12,["Item","2026E","2027E","2028E","2029E","2030E"])
fc=[]
for i in range(5):
    r0=RP[i];n=round(r0*EM[i]*(1-TR),1);d=round(r0*DR,1);c2=round(-r0*CX[i],1)
    dd=r0-(RH[4] if i==0 else RP[i-1]);w=round(-dd*NW[i],1);u=round(n+d+c2+w,1);fc.append(u)
di=[("Rev",[round(v,1) for v in RP]),("EBIT",[round(RP[i]*EM[i],1) for i in range(5)]),
    ("Tax",[round(RP[i]*EM[i]*TR,1) for i in range(5)]),("NOPAT",[round(RP[i]*EM[i]*(1-TR),1) for i in range(5)]),
    ("+DA",[round(RP[i]*DR,1) for i in range(5)]),("-Capex",[round(-RP[i]*CX[i],1) for i in range(5)]),
    ("-WCchg",[round(-(RP[i]-(RH[4] if i==0 else RP[i-1]))*NW[i],1) for i in range(5)]),
    ("=UFCF",fc)]
for idx,(lb,vs) in enumerate(di):
    sw(dc,13+idx,1,lb)
    for j in range(5):sw(dc,13+idx,j+2,vs[j],fill=OF if idx==7 else None,nf="#,##0")
sw(dc,22,1,"Discount",HFR,HF);hr(dc,23,["Item","2026E","2027E","2028E","2029E","2030E","TV"])
sw(dc,24,1,"FCF")
for i in range(5):sw(dc,24,i+2,fc[i],nf="#,##0")
sw(dc,25,1,"Year")
for i in range(5):sw(dc,25,i+2,i+0.5,nf="0.0")
sw(dc,26,1,"DF")
pv=[]
for i in range(5):df=1/(1+WA)**(i+0.5);pv.append(round(fc[i]*df,1));sw(dc,26,i+2,round(df,4),nf="0.0000")
sw(dc,27,1,"PV")
for i in range(5):sw(dc,27,i+2,pv[i],fill=OF,nf="#,##0")
tf=round(fc[-1]*(1+TG),1);tv=round(tf/(WA-TG),1);pt=round(tv/(1+WA)**4.5,1)
sw(dc,29,1,"TV");sw(dc,30,1,"TermFCF");sw(dc,30,7,tf,nf="#,##0");sw(dc,31,1,"TV");sw(dc,31,7,tv,nf="#,##0");sw(dc,32,1,"PVTV");sw(dc,32,7,pt,fill=OF,nf="#,##0")
sp=round(sum(pv),1);ev=round(sp+pt,1);eq=round(ev+NC,1);ip=round(eq/SH,2);ud=round(ip/CP-1,4)
sw(dc,34,1,"Summary",HFR,HF);sw(dc,35,1,"SumFCF");sw(dc,35,7,sp,nf="#,##0")
sw(dc,36,1,"PVTV");sw(dc,36,7,pt,nf="#,##0");sw(dc,37,1,"EV");sw(dc,37,7,ev,fill=OF,nf="#,##0")
sw(dc,38,1,"+Cash");sw(dc,38,7,NC,nf="#,##0");sw(dc,39,1,"Eq");sw(dc,39,7,eq,fill=OF,nf="#,##0")
sw(dc,40,1,"Shares");sw(dc,40,7,SH,nf="#,##0");sw(dc,41,1,"Price");sw(dc,41,7,ip,fill=OF,nf="0.00")
sw(dc,42,1,"Curr");sw(dc,42,7,CP,nf="0.00");sw(dc,43,1,"Up/Down");sw(dc,43,7,ud,nf="0.0%")
sw(dc,44,1,"TV%EV");sw(dc,44,7,round(pt/ev,4),nf="0.0%")
sw(dc,46,1,"WACCDetail",HFR,HF)
wi=[("Rf",0.0176),("Beta",0.60),("ERP",0.06),("Ke",0.0176+0.60*0.06),("Kd",0.04),("Tax",0.15),("Kd_at",0.04*(1-0.15)),("Mcap",265.91),("NetD",-3.39),("EV",262.52),("Ewt",265.91/262.52),("Dwt",-3.39/262.52),("WACC",0.0176+0.60*0.06)]
for idx,(lb,val) in enumerate(wi):sw(dc,47+idx,1,lb);sw(dc,47+idx,2,val,BF,IF,"0.00%" if isinstance(val,float) else "0.00")
sw(dc,62,1,"Sensitivity:Price/WACC/TermG",HFR,HF)
wr=[0.044,0.049,0.054,0.059,0.064];tr=[0.020,0.025,0.030,0.035,0.040]
sw(dc,63,1,"WACC\\TG")
for j,tgv in enumerate(tr):sw(dc,63,j+2,tgv,nf="0.0%")
for i,waccv in enumerate(wr):
    sw(dc,64+i,1,waccv,nf="0.0%")
    for j,tgv in enumerate(tr):
        tf2=fc[-1]*(1+tgv);tv2=tf2/(waccv-tgv);pt2=tv2/(1+waccv)**4.5
        ps2=sum(fc[k]/(1+waccv)**(k+0.5) for k in range(5));ev2=ps2+pt2;eq2=ev2+NC;ip2=eq2/SH
        sw(dc,64+i,j+2,round(ip2,2),nf="0.00",fill=PatternFill("solid",fgColor="BDD7EE") if(i==2 and j==2)else None)
dc.column_dimensions['A'].width=22
for c in range(2,9):dc.column_dimensions[get_column_letter(c)].width=16
cp=wb.create_sheet("Comps");cp.merge_cells("A1:G1")
sw(cp,1,1,"MoonsElectric Comps",TF);hr(cp,3,["Co","Ticker","Mcap","Rev","PE_TTM","PE_26E","PB"])
cm=[["Moons","603728.SH",265.91,27.62,392.4,258,8.94],["LeiSai","002979.SZ",138.0,18.74,57.0,43,8.5],["Inovance","300124.SZ",2030.0,350.0,43.4,28,8.0],["Veichi","688698.SH",146.0,12.5,56.9,45,6.5]]
for ic,c in enumerate(cm):
    r=4+ic;inp=(ic==0)
    for jc,v in enumerate(c):
        f=BF if inp else BK;fl=IF if inp else None
        if jc==1:sw(cp,r,jc+1,v,f,fl)
        elif isinstance(v,(int,float)):sw(cp,r,jc+1,v,f,fl,"0.00" if jc>=4 else "#,##0.0")
        else:sw(cp,r,jc+1,v,f,fl)
pv2=[57.0,43.4,56.9];p26=[43,28,45];pb2=[8.5,8.0,6.5]
mp=round(sorted(pv2)[1],1);m26=sorted(p26)[1];mpb=round(sorted(pb2)[1],1)
ap2=round(sum(pv2)/3,1);a26=round(sum(p26)/3,1);apb=round(sum(pb2)/3,1)
sw(cp,9,1,"Median_exMoons",BD);sw(cp,9,5,mp,nf="0.0");sw(cp,9,6,m26,nf="0.0");sw(cp,9,7,mpb,nf="0.0")
sw(cp,10,1,"Mean_exMoons",BD);sw(cp,10,5,ap2,nf="0.0");sw(cp,10,6,a26,nf="0.0");sw(cp,10,7,apb,nf="0.0")
en=125
sw(cp,12,1,"PEValuation",HFR,HF);sw(cp,13,1,"2026E_NP_RMBm");sw(cp,13,2,en,BF,IF,"#,##0")
sw(cp,14,1,"PeerPE26");sw(cp,14,2,m26,BF,IF,"0.0x")
im=round(en*m26/100,1);sw(cp,15,1,"ImplMcap");sw(cp,15,2,im,nf="#,##0.0")
ip2=round(im*100/SH,2);sw(cp,16,1,"ImplPrice");sw(cp,16,2,ip2,nf="0.00")
sw(cp,17,1,"CurrPrice");sw(cp,17,2,CP,nf="0.00")
sw(cp,18,1,"UpDown");sw(cp,18,2,round(ip2/CP-1,4),nf="0.0%")
for c in range(1,9):cp.column_dimensions[get_column_letter(c)].width=18
out=r"d:\claude-for-finance\outputs\鸣志电器\鸣志电器_三表模型_DCF估值.xlsx"
wb.save(out)
print(f"OK,DCF={ip},PE={ip2}")
