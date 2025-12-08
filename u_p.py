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

    # 阻尼考慮下的穩態響應 q_n (SDOF 公式)
    driver_omega = 6.8 * np.pi  # 激振角頻率 (rad/s)
    zeta_list = [0.006, 0.00493, 0.006]  # 阻尼比(對應三個模態)
    zeta = zeta_list[i]

    beta = driver_omega / omega[i]
    phi_phase = np.arctan2(2 * zeta * beta, (1 - beta ** 2))  # 相位差 Phi

    # 荷載幅值 P0 (N)
    P0 = 0.0545 * 0.02 * (6.8 * np.pi) ** 2

    # q_n(t) = A * sin(Ω t - Phi)，其中 A = P0/k_n * 1/ sqrt((1-β^2)^2 + (2βζ)^2)
    ampl = (P0 / modal_stiffness) / np.sqrt((1 - beta ** 2) ** 2 + (2 * beta * zeta) ** 2)

    print("\n--- 阻尼穩態響應參數 (模態座標 q_n) ---")
    print(f"阻尼比 ζ = {zeta:.5f}")
    print(f"β = Ω/ω_n = {beta:.5f}")
    print(f"相位差 Phi = {phi_phase:.5f} rad")
    print(f"振幅 A = {ampl:.6e} (同 q_n 單位)")
    print("q_n(t) = A * sin(Ω*t - Phi)")
    
    # 驗證: 計算 (k - x*m)φ 應該接近 0
    verification = A @ phi
    print(f"\n驗證 [k - ω²m]φ = (應接近0)")
    print(verification)
    print(f"誤差範數: {np.linalg.norm(verification):.2e}")