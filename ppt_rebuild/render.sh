#!/bin/bash
# render.sh FIRST LAST : build those slides, render via LibreOffice, make side-by-side compare PNGs
S=/tmp/claude-0/-home-user-bmk/ec77546d-a9b0-5b71-9d6c-3073ce8e76bd/scratchpad
export SLIDE_IMG_DIR=$S/img
cd "$(dirname "$0")"
python3 build.py $S/t.pptx $1 $2 >/dev/null || exit 1
rm -rf $S/r && mkdir -p $S/r
(cd $S/r && soffice --headless --convert-to pdf ../t.pptx >/dev/null 2>&1 && pdftoppm -r 96 -png t.pdf p >/dev/null 2>&1)
python3 - "$1" "$2" <<'PY'
import sys, glob, os
from PIL import Image
S='/tmp/claude-0/-home-user-bmk/ec77546d-a9b0-5b71-9d6c-3073ce8e76bd/scratchpad'
a,b=int(sys.argv[1]),int(sys.argv[2])
pages=sorted(glob.glob(S+'/r/p-*.png'))
nums=[n for n in range(a,b+1) if os.path.exists('slides/s%03d.py'%n)]
for n,p in zip(nums,pages):
    o=Image.open(S+'/img/s%03d.png'%n).convert('RGB'); r=Image.open(p).convert('RGB').resize((960,720))
    c=Image.new('RGB',(1930,720),'red'); c.paste(o,(0,0)); c.paste(r,(970,0)); c.save(S+'/r/cmp%03d.png'%n)
print('ok',len(pages))
PY
