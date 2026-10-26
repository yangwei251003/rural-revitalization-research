"""Render an explicitly labelled chart from a completed synthetic analysis run."""
from pathlib import Path
import argparse,json
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

def main():
 p=argparse.ArgumentParser();p.add_argument("--run",type=Path,required=True);p.add_argument("--out",type=Path,required=True);a=p.parse_args()
 manifest=json.loads((a.run/"run-manifest.json").read_text(encoding="utf-8"))
 if manifest["dataIdentity"]!="synthetic-schema-demonstration-not-observed-data":p.error("Only explicitly synthetic runs may be published by this renderer")
 x=a.run/"Step1_熵权TOPSIS结果.xlsx";trend=pd.read_excel(x,"时序统计描述");weights=pd.read_excel(x,"指标权重")
 plt.rcParams.update({"font.family":"DejaVu Sans","axes.spines.top":False,"axes.spines.right":False,"axes.labelcolor":"#bdcce0","text.color":"#e4edf9","xtick.color":"#a6bbd3","ytick.color":"#a6bbd3","axes.facecolor":"#111f32","figure.facecolor":"#091423","axes.edgecolor":"#476078"})
 fig,axes=plt.subplots(1,2,figsize=(14,5.4),gridspec_kw={"width_ratios":[1.2,1]})
 axes[0].plot(trend["年份"],trend["全国均值"],"o-",color="#61ddc9",linewidth=2.8);axes[0].set_title("Mean score from synthetic inputs",loc="left",pad=20,fontweight="bold");axes[0].set_xlabel("Schema year");axes[0].set_ylabel("TOPSIS score");axes[0].grid(alpha=.12)
 codes=weights["指标代码"].str.slice(0,3);axes[1].barh(codes,weights["权重(%)"],color="#7aa9ff");axes[1].set_title("Entropy weights from the same demo",loc="left",pad=20,fontweight="bold");axes[1].set_xlabel("Weight (%)")
 fig.suptitle("SYNTHETIC SCHEMA DEMONSTRATION",fontsize=18,fontweight="bold",y=.98)
 fig.text(.5,.025,"NOT OBSERVED DATA  |  31 schema labels × 8 years  |  Seed 20261026  |  No regional or policy inference",ha="center",fontsize=10,color="#61ddc9")
 fig.tight_layout(rect=[0,.06,1,.91]);a.out.mkdir(parents=True,exist_ok=True);fig.savefig(a.out/"demo-summary.png",dpi=150);plt.close(fig)
 summary={"identity":"SYNTHETIC DEMO — NOT OBSERVED DATA","rows":248,"years":[2015,2022],"indicators":14,"seed":20261026,"meanTrend":[{"year":int(y),"score":float(v)} for y,v in zip(trend["年份"],trend["全国均值"])],"analysisStatus":manifest["status"],"scientificValidityAssessed":False}
 (a.out/"demo-summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
if __name__=="__main__":main()
