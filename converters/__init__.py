from converters.layout1 import convert_file as convert_layout1
from converters.layout2 import convert_file as convert_layout2
from converters.layout3 import convert_file as convert_layout3
from converters.layout4 import convert_file as convert_layout4

CONVERTERS = {
    "layout1": convert_layout1,
    "layout2": convert_layout2,
    "layout3": convert_layout3,
    "layout4": convert_layout4,
}
