"""Synthetic data with the legacy schema; no real provincial observations."""
from pathlib import Path
import argparse
import ast
import numpy as np
import pandas as pd

ROOT=Path(__file__).resolve().parents[1]


def constants(filename,name):
 tree=ast.parse((ROOT/'analysis/legacy'/filename).read_text(encoding='utf8'))
 for node in tree.body:
  if isinstance(node,ast.Assign) and any(isinstance(t,ast.Name) and t.id==name for t in node.targets):
   return ast.literal_eval(node.value)
 raise KeyError(name)


def main():
 parser=argparse.ArgumentParser();parser.add_argument('--out',type=Path,required=True);args=parser.parse_args()
 if args.out.exists():parser.error('拒绝覆盖已有面板')
 provinces=constants('build_panel.py','provinces_31')
 columns=constants('step1_entropy_topsis.py','positive')+constants('step1_entropy_topsis.py','negative')
 rng=np.random.default_rng(20261026)
 scales={'X01':15000,'X02':3500,'X03':2000,'X04':50,'X05':0.08,'X06':0.008,'X07':12,'X08':700,'X09':90,'X11':35,'X14':15,'X15':5000,'X16':900,'X17':1600}
 rows=[]
 for i,province in enumerate(provinces):
  for year in range(2015,2023):
   row={'省份':province,'年份':year}
   for column in columns:
    # Varied independent components exercise numerical code, without encoding research findings.
    row[column]=scales[column[:3]]*(0.6+0.6*rng.random()+0.015*(year-2015)+0.1*np.sin(i))
   rows.append(row)
 panel=pd.DataFrame(rows)
 with pd.ExcelWriter(args.out,engine='openpyxl') as writer:
  panel.to_excel(writer,sheet_name='面板数据',index=False)
  pd.DataFrame({'说明':['SYNTHETIC DEMO — NOT OBSERVED DATA','31个真实行政区名称仅用于匹配原脚本结构，数值全部由固定随机种子生成。','不是国家统计局数据，不用于省份排名、政策判断、统计结论或因果推断。'],'seed':[20261026,None,None]}).to_excel(writer,sheet_name='合成身份',index=False)
 print('已生成248行完全合成的结构演示面板。')


if __name__=='__main__':main()
