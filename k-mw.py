import numpy as np
from scipy.linalg import eig

# 定義矩陣k和m
k = 58742.57326* np.array([[2, -1, 0],
                       [-1, 2, -1],
                       [0, -1, 1]], dtype=float)
m = np.array([[7.5375, 0, 0],
              [0, 7.5375, 0],
              [0, 0, 8.6275]], dtype=float)

# 使用 scipy.linalg.eig 求解廣義特徵值問題
# 求解 k*v = λ*m*v，等價於 det(k - λm) = 0
eigenvalues, eigenvectors = eig(k, m)

# 只取實數根（虛部非常小的視為實數）
eigenvalues = np.real(eigenvalues[np.abs(np.imag(eigenvalues)) < 1e-10])
eigenvalues = np.sort(eigenvalues)

print("The solutions x for |k - mx| = 0 are:", eigenvalues)
print("\n角頻率 ω (omega) = sqrt(x):")
omega = np.sqrt(eigenvalues)
print(omega)
print("\n自然頻率 f (Hz) = ω/(2π):")
frequency = omega / (2 * np.pi)
print(frequency)
