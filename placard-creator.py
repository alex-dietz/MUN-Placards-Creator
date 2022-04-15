import copy
import json
from pptx import Presentation


pres = Presentation("test.pptx")

def duplicate_slide(pres, index):
    template = pres.slides[index]
    try:
        blank_slide_layout = pres.slide_layouts[12]
    except:
        blank_slide_layout = pres.slide_layouts[len(pres.slide_layouts)-1]

    copied_slide = pres.slides.add_slide(blank_slide_layout)

    for shp in template.shapes:
        el = shp.element
        newel = copy.deepcopy(el)
        copied_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')   

    return copied_slide

def find_shape_by_name(shapes, name):
    for shape in shapes:
        if shape.name == name:
            return shape
    return None

def add_text(shape, text, alignment=None):

    if alignment:
        shape.vertical_anchor = alignment

    tf = shape.text_frame
    tf.clear()
    run = tf.paragraphs[0].add_run()
    run.text = text if text else ''



copy = duplicate_slide(pres,0)
""" slide_title = find_shape_by_name(slide.shapes,'slide_title')
add_text(slide_title,'TEST SLIDE') """

pres.save('test.pptx')