import glob
from ppt_merger import ppt_merger


input_files = glob.glob(r'input\RP2021*.pptx')
ppt_merger(input_files=input_files,
           section_names=[fname.replace('input\\RP', '').replace('.pptx', '') for fname in input_files],
           output_fname=r'output\RP2021.pptx')

input_files = glob.glob(r'input\RP2022*.pptx')
ppt_merger(input_files=input_files,
           section_names=[fname.replace('input\\RP', '').replace('.pptx', '') for fname in input_files],
           output_fname=r'output\RP2022.pptx')
