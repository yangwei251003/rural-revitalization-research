"""Build the dependency-free static demo; no private inputs are copied."""
from pathlib import Path
import argparse,shutil
root=Path(__file__).resolve().parents[1]
p=argparse.ArgumentParser();p.add_argument("--out",type=Path,default=Path("_site"));a=p.parse_args();out=a.out.resolve()
if out.exists() and any(out.iterdir()):p.error("Refusing to overwrite a nonempty site directory")
out.mkdir(parents=True,exist_ok=True);(out/"images").mkdir(exist_ok=True)
shutil.copyfile(root/"docs/preview/index.html",out/"index.html")
shutil.copyfile(root/"docs/images/demo-summary.png",out/"images/demo-summary.png")
(out/".nojekyll").touch();print("Built explicitly synthetic static preview")
