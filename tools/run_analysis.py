"""Run archived analysis scripts in a fresh workspace; never overwrite historical results."""
from pathlib import Path
import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime, timezone

ROOT=Path(__file__).resolve().parents[1]
PANEL_NAME='乡村振兴面板数据_2015-2022.xlsx'
STEPS={
 '1':'step1_entropy_topsis.py','2':'step2_spatial_temporal.py','3':'step3_dagum_gini.py',
 '4':'step4_panel_regression.py','5':'step5_robustness.py','6':'step6_heterogeneity_regression.py',
 '7a':'step7_1_robustness_alternative_dv.py','7b':'step7_4_robustness_alternative_method.py',
 '8':'Step8_空间计量分析.py','9':'Step9_收敛性分析.py','10':'Step10_障碍度诊断.py',
}


def sha(path):
 return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def main():
 parser=argparse.ArgumentParser(description='独立复现乡村振兴分析步骤，默认1–3')
 source=parser.add_mutually_exclusive_group(required=True)
 source.add_argument('--demo',action='store_true',help='生成完全合成的同结构面板')
 source.add_argument('--panel',type=Path,help='使用者自备且有权使用的面板工作簿')
 parser.add_argument('--out',type=Path,required=True,help='新建或空输出目录')
 parser.add_argument('--steps',default='1,2,3',help='逗号分隔，或all；可选1..10、7a、7b')
 args=parser.parse_args()
 output=args.out.resolve()
 if output.exists() and any(output.iterdir()):parser.error('输出目录非空，拒绝覆盖既有结果')
 output.mkdir(parents=True,exist_ok=True)
 for name in ['02_分析结果','03_图表文件','logs']:(output/name).mkdir(exist_ok=True)
 if args.demo:
  subprocess.run([sys.executable,str(ROOT/'tools/make_demo_panel.py'),'--out',str(output/PANEL_NAME)],check=True)
  identity='synthetic-schema-demonstration-not-observed-data'
 else:
  if not args.panel.is_file():parser.error('输入面板不存在')
  shutil.copyfile(args.panel,output/PANEL_NAME)
  identity='user-supplied-local-data-not-cleared-for-publication'
 selected=list(STEPS) if args.steps=='all' else args.steps.split(',')
 if any(step not in STEPS for step in selected):parser.error('未知步骤')
 manifest={'createdAtUTC':datetime.now(timezone.utc).isoformat(),'dataIdentity':identity,
           'inputSha256':sha(output/PANEL_NAME),'steps':[],'scientificValidityAssessed':False,
           'notice':'软件运行状态不等于数据、方法、因果效应或论文结论得到新验证。'}
 env={**os.environ,'PYTHONIOENCODING':'utf-8','MPLBACKEND':'Agg','RESEARCH_WORKSPACE':str(output),
      'OMP_NUM_THREADS':'1','OPENBLAS_NUM_THREADS':'1'}
 failed=False
 for step in selected:
  script=ROOT/'analysis/legacy'/STEPS[step]
  with (output/'logs'/f'step-{step}.log').open('w',encoding='utf8') as log:
   result=subprocess.run([sys.executable,str(script)],cwd=output,env=env,stdout=log,stderr=subprocess.STDOUT)
  manifest['steps'].append({'id':step,'script':STEPS[step],'scriptSha256':sha(script),'exitCode':result.returncode,
                            'status':'PASS' if result.returncode==0 else 'FAIL'})
  print(f'Step {step}: {manifest["steps"][-1]["status"]}',flush=True)
  if result.returncode:
   failed=True;break
  if step=='1':shutil.copyfile(output/'Step1_熵权TOPSIS结果.xlsx',output/'02_分析结果/Step1_熵权TOPSIS结果.xlsx')
 manifest['status']='FAIL' if failed else 'PASS'
 (output/'run-manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
 if failed:raise SystemExit('步骤失败，已保留原日志；未将失败结果包装成完成。')
 print('运行完成。历史研究文件未被改动。',flush=True)


if __name__=='__main__':main()
