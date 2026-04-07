import pandas as pd
import ofxparse
import os
from datetime import datetime

df = pd.DataFrame()
for extrato in os.listdir("extratos"):
    print(extratos)