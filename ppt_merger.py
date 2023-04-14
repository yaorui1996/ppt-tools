import os
import glob
import numpy as np
from win32com.client import Dispatch


def ppt_merger(input_files, section_names=None, output_fname=r'output\output.pptx'):
    ppt = Dispatch('PowerPoint.Application')
    p0 = ppt.Presentations.Add()

    slides_counts = []
    for pi_fname in input_files:
        pi = ppt.Presentations.Open(os.path.join(os.getcwd(), pi_fname))
        slides_counts.append(pi.Slides.Count)
        for i in range(pi.Slides.Count):
            pi.Slides(i+1).Copy()
            p0.Slides.Paste()
        pi.Close()
    if section_names is not None:
        indexes = np.cumsum([1, *slides_counts[: -1]])
        for i in range(len(input_files)):
            # p0.SectionProperties.AddBeforeSlide(SlideIndex=indexes[i], sectionName=input_files[i].replace('input\\RP', '').replace('.pptx', ''))
            p0.SectionProperties.AddBeforeSlide(SlideIndex=indexes[i], sectionName=section_names[i])

    # p0.SaveAs(os.path.join(os.getcwd(), r'output\output.pptx'))
    p0.SaveAs(os.path.join(os.getcwd(), output_fname))
    p0.Close()
    ppt.Quit()
