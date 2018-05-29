import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import itertools

# Fixing random state for reproducibility
np.random.seed(19680801)


matplotlib.rcParams['axes.unicode_minus'] = False
fig, ax = plt.subplots()
ax.plot(10*np.random.randn(100), 10*np.random.randn(100), 'o')
ax.set_title('Using hyphen instead of Unicode minus')
plt.show()

print(list(itertools.product([0,1], repeat=3)))