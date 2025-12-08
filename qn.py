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

print("特徵值 x (ω²) = ", eigenvalues)
print("\n角頻率 ω (omega) = sqrt(x):")
omega = np.sqrt(eigenvalues)
print(omega)
print("\n自然頻率 f (Hz) = ω/(2π):")
frequency = omega / (2 * np.pi)
print(frequency)

print("\n" + "="*60)
print("計算每個特徵值對應的振型 φ (phi)")
print("="*60)

# 對每個特徵值計算對應的特徵向量
for i, x_val in enumerate(eigenvalues):
    print(f"\n第 {i+1} 個模態 (x = {x_val:.4f}):")
    print(f"角頻率 ω = {omega[i]:.4f} rad/s")
    print(f"自然頻率 f = {frequency[i]:.4f} Hz")
    
    # 計算 k - x*m
    A = k - x_val * m
    print(f"\n矩陣 [k - ω²m] =")
    print(A)
    
    # 求解 (k - x*m)φ = 0 的非零解
    # 使用特徵向量(已經從 eig 函數得到)
    phi = eigenvectors[:, i]
    
    # 歸一化: 將最大值設為 1
    phi_normalized = phi / np.max(np.abs(phi))
    
    print(f"\n特徵向量 φ{i+1} (歸一化) =")
    print(phi_normalized)

    # 模態質量與剛度 (phi^T M phi, phi^T K phi)
    modal_mass = np.conjugate(phi_normalized).T @ m @ phi_normalized
    modal_stiffness = np.conjugate(phi_normalized).T @ k @ phi_normalized
    print("\n模態質量 phi^T M phi =", np.real_if_close(modal_mass))
    print("模態剛度 phi^T K phi =", np.real_if_close(modal_stiffness))
    
    # 驗證: 計算 (k - x*m)φ 應該接近 0
    verification = A @ phi
    print(f"\n驗證 [k - ω²m]φ = (應接近0)")
    print(verification)
    print(f"誤差範數: {np.linalg.norm(verification):.2e}")
