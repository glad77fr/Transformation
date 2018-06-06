import pandas as pd

a = pd.DataFrame({"Mat": [1,4,6,7,3], 2: [1,4,5,5,5], 6: [2,4,5,5,0]})
b = pd.DataFrame({"Mat": [1,4,4,5,5], 2: [11,144,5,5,65], 4: [12,41,51,51,11]})

c = pd.merge(a,b,on="Mat",how='left')
print(a)
print(b)
print(c)