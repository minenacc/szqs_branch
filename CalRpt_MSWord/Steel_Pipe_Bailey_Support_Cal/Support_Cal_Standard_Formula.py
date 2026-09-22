
# 钢结构设计标准GB50017_2017 公式 8.2.1-1/2
GB50017_2017_8_2_1_1 = r'\sigma=\frac{N}{\varphi A}+\frac{\beta M}{\gamma_mW(1-0.8\frac{N}{N_{Ex}^\prime})}' # NEX
GB50017_2017_8_2_1_2 = r'N_{EX}^\prime=\frac{\pi^2EA}{1.1\ \lambda^2}' # 平面稳定性计算应力结果

# 建筑地基基础设计规范GB50007_2011 公式 5.2.2-1
GB50007_2011_5_2_2_1 = r'p_k=\frac{\left(F_k+G_k\right)}{A}' # pk

# 建筑地基基础设计规范GB50007_2011 条文 5.2.1.3 计算偏心距
GB50007_2011_Eccentricity = r'e=\frac{M_k}{\left(F_k+G_k\right)}' # e

# 建筑地基基础设计规范GB50007_2011 公式 5.2.2-2/3
GB50007_2011_5_2_2_2 = r'p_{kmax}=\frac{\left(F_k+G_k\right)}{A}+\frac{M_k}{W}' # pkmax(小偏心)
GB50007_2011_5_2_2_3 = r'p_{kmin}=\frac{\left(F_k+G_k\right)}{A}-\frac{M_k}{W}' # pkmin(小偏心)

# 建筑地基基础设计规范GB50007_2011 公式 5.2.2-4
GB50007_2011_5_2_2_4 = r'p_{kmax}=\frac{2\left(F_k+G_k\right)}{3la}' # pkmax(大偏心)

# 建筑地基基础设计规范GB50007_2011 公式 5.2.7-3
GB50007_2011_5_2_7_3 = r'p_z=\frac{b\left(p_k-p_c\right)}{\left(b+2ztan\theta\right)\left(l+2ztan\theta\right)}' # pz

# 建筑地基基础设计规范GB50007_2011 条文 8.2.8 式中pj
GB50007_2011_5_2_8_pj = r'p_j=K_s\left(p_k-\frac{G_k}{A}\right)'

# 建筑地基基础设计规范GB50007_2011 条文 8.2.11 式中pmax , 公式中pmax是总反力, 计算结果等同于净反力代替
GB50007_2011_5_2_11_pjmax = r'p_{jmax}=K_s\left(p_{kmax}-\frac{G_k}{A}\right)'

# 建筑地基基础设计规范GB50007_2011 条文 8.2.11 式中pmin, 公式中pmin是总反力, 计算结果等同于净反力代替
GB50007_2011_5_2_11_pjmin = r'p_{jmin}=K_s\left(p_{kmin}-\frac{G_k}{A}\right)'

# 建筑桩基技术规范JGJ94-2008 公式 5.3.7-1
JGJ94_2008_5_3_7_1 = r'Q_{uk}=Q_{sk}+Q_{pk}=\mu\sum{q_{sik}l_i+\lambda_pq_{pk}A_p}' # 钢管桩Quk
# 建筑桩基技术规范JGJ94-2008 公式 5.3.8-1
JGJ94_2008_5_3_8_1 = r'Q_{uk}=Q_{sk}+Q_{pk}=\mu\sum{q_{sik}l_i+q_{pk}\left(A_j+\lambda_pA_{pl}\right)}' # 预制桩Quk
# 建筑桩基技术规范JGJ94-2008 公式 5.3.6
JGJ94_2008_5_3_6 = r'Q_{uk}=Q_{sk}+Q_{pk}=\mu\sum{\psi_{si}q_{sik}l_i+\psi_pq_{pk}A_p}' # 灌注桩Quk
# 建筑桩基技术规范JGJ94-2008 公式 5.2.2
JGJ94_2008_5_2_2 = r'R_a=\frac{1}{K}Q_{uk}' # Ra

# 建筑桩基技术规范JGJ94-2008 公式 5.7.2-1 式中W0
JGJ94_2008_5_7_2_1_W0 = r'W_0=\frac{\pi d}{32}\left[d^2+2\left(\alpha_E-1\right)\rho_gd_0^2\right]'

# 建筑桩基技术规范JGJ94-2008 公式 5.7.2-1 式中An
JGJ94_2008_5_7_2_1_An = r'A_n=\frac{\pi d^2}{4}\left[1+\left(\alpha_E-1\right)\rho_g\right]'

# 建筑桩基技术规范JGJ94-2008 公式 5.7.2-1
JGJ94_2008_5_7_2_1 = r'R_{ha}=\frac{0.75\alpha\gamma_mf_tW_0}{\nu_M}\left(1.25+22\rho_g\right)\left(1\pm\frac{\zeta_N\bullet N}{\gamma_mf_tA_n}\right)' # Rha

# 建筑基坑支护技术规程JGJ 120-2012 公式 4.1.6
JGJ_120_2012_4_1_6 = r'm=\frac{0.2\varphi^2-\varphi+c}{v_b}' # m

# 建筑桩基技术规范JGJ94-2008 公式 5.7.5
JGJ94_2008_5_7_5 = r'\alpha=\sqrt[5]{\frac{mb_0}{EI}}' # alpha

# 建筑桩基技术规范JGJ94-2008 公式 5.7.2-2
JGJ94_2008_5_7_2_2 = r'R_{ha}=0.75\frac{\alpha^3EI}{v_x}x_{0a}' # Rha

 # 公路桥梁抗风设计规范
JTG_3360_01_2018_4_2_4 = r'U_{s10}=k_cU_{10}'
JTG_3360_01_2018_4_2_6_1 = r'U_d=k_f\left(\frac{Z}{10}\right)^{\alpha_0}U_{s10}'
JTG_3360_01_2018_4_2_6_2 = r'U_d=k_fk_tk_hU_{10}'
JTG_3360_01_2018_5_2_1 = r'U_g=G_VU_d'
JTG_3360_01_2018_5_3_1 = r'F_g/D=\frac{1}{2}\rho U_g^2C_H'

# 港口工程荷载规范
JTS_144_1_2010_13_0_1 = r'F_\omega=C_\omega\frac{\rho}{2}V^2A'