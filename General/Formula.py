# 1. 标准库
import math
from operator import mul

# 3. 本地模块
from General.DataUtils import double_param_table_index


# 计算主动土压力系数
def cal_Ka(fai):
    Ka = round((math.tan(math.radians(45)-math.radians(fai/2)))**2, 3)
    return Ka

# 计算被动土压力系数
def cal_Kp(fai):
    Kp = round((math.tan(math.radians(45)+math.radians(fai/2)))**2, 3)
    return Kp

# 计算m值
def cal_m(fai, c):
    m = round((0.2*pow(fai,2)-fai+c)/10, 3)*1000
    return m

# 计算等代土弹簧刚度
def cal_Ks(m , z, a=1, b=1):
    Ks = round(a * b * m * abs(z), 1)
    return Ks

# 计算地面均布荷载作用下、深度为h的的土体竖向应力标准值_土 # σ0+(γ土-γ水)*h
def cal_sigmak_solid(sigma1, r1, h1, h2, r2=9.8):
    sigma2 = sigma1 + (r1-r2)*abs(h2-h1) # sigma1为土的层顶竖向应力标准值, h2为层底标高, h1为层顶标高, r为当前层土的饱和容重
    return [sigma1, sigma2]

# 计算地面均布荷载作用下、深度为h的的土体竖向应力标准值_水 # γ水h
def cal_sigmak_water(sigma1, h1, h2, r2=9.8):
    sigma2 = sigma1 + r2*abs(h2-h1) # sigma1为水的层顶竖向应力标准值, h2为层底标高, h1为层顶标高, r为当前层土的饱和容重
    return [sigma1, sigma2]

# 计算主动土压力_土
def cal_pak_soild(Ka, sigma1, sigma2, c):
    # sigma1, sigma2 = cal_sigmak_solid(sigma1, r1, h1, h2) # σ0+(γ土-γ水)*h
    pak1_soild = sigma1*Ka-2*c*pow(Ka,0.5)
    pak2_soild = sigma2*Ka-2*c*pow(Ka,0.5) # (σ0+(γ土-γ水)*h)*Ka-2cKa^0.5
    return [pak1_soild, pak2_soild]

# 计算主动土压力_水
def cal_pak_water(Ka, sigma1, sigma2, method):
    if method == '水土合算':
        a = 1
    elif method == '水土分算':
        a = 0
    else:
        a = 0  # 默认水土分算
    pak1_water = sigma1 * pow(Ka, a)
    pak2_water = sigma2 * pow(Ka, a)
    return [pak1_water, pak2_water]


# 计算被动土压力_土
def cal_ppk_soild(Kp, sigma1, sigma2, c):
    # sigma1, sigma2 = cal_sigmak_solid(sigma1, r1, h1, h2) # σ0+(γ土-γ水)*h
    ppk1_soild = sigma1*Kp+2*c*pow(Kp,0.5)
    ppk2_soild = sigma2*Kp+2*c*pow(Kp,0.5) # (σ0+(γ土-γ水)*h)*Kp+2cKp^0.5
    return [ppk1_soild, ppk2_soild]


# 计算被动土压力_水
def cal_ppk_water(Kp, sigma1, sigma2, method):
    if method == '水土合算':
        a = 1
    elif method == '水土分算':
        a = 0
    else:
            a = 0  # 默认水土分算
    # sigma1, sigma2 = cal_sigmak_water(sigma1, h1, h2) # γ水h
    ppk1_water = sigma1*pow(Kp,a)
    ppk2_water = sigma2*pow(Kp,a) # γ水h*Ka, 分算时Ka相当于=1
    return [ppk1_water, ppk2_water]


# 计算分布土反力_土
def cal_ps0_soild(Ka, sigma1, sigma2):
    ps01_soild = sigma1*Ka
    ps02_soild = sigma2*Ka # (σ0+(γ土-γ水)*h)*Ka
    return [ps01_soild, ps02_soild]


# 计算土弹簧时圆形桩的换算宽度
# 根据《建筑基坑支护技术规程》4.1.7条计算
def cal_width_trans(D, Gap, Type = 'Circle'):
    '''
    Putin:
    Type 类型
    D    直径或宽度, 单位m
    Gap  间距, 单位m
    Return:
    b0   换段宽度, 单位m
    '''
    n = 0.9 if Type == 'Circle' else 1
    if D<=1:
        b1 = n*(1.5*D+0.5)
    else:
        b1 = n*(D+1)
    b0 = b1 if b1 <= Gap else Gap
    return b0


# 矩形局部附加荷载计算函数
def Localized_Rectangular_Load(angle, p0, b, a, d, l):
    zamin = d + (a / math.tan(angle))
    zamax = d + ((3 * a + b) / math.tan(angle))
    sigma_k = (p0 * b * l) / ((b + 2 * a) * (l + 2 * a))
    return [zamin, zamax, sigma_k]


# 条形局部附加荷载计算函数
def Localized_Strip_Load(angle, p0, b, a, d):
    zamin = d + (a / math.tan(angle))
    zamax = d + ((3 * a + b) / math.tan(angle))
    sigma_k = (p0 * b) / (b + 2 * a)
    return [zamin, zamax, sigma_k]


class Flow_JTS_144_1_2010:

    # 基本参数，环境ENV，设计流速V，水位高程H
    def __init__(self, ENV, V, Hdelta):
        # 淡水or海水环境，影响水密度
        self.ENV = float(ENV)
        # 设计流速
        self.V = float(V)
        # 水位高程，结合桩底高程确定钢管桩受水流影响长度
        self.Hdelta = float(Hdelta)
        # 后墩的遮流影响系数m1
        self.m1_dict = {"L/D": [[1,2,3,4,6,8,12,16,18,20], [-0.38,0.25,0.54,0.66,0.78,0.82,0.86,0.88,0.90,1.00]]}
        # 淹没深度影响系数n1 = 1.0
        self.n1 = 1.0
        # 墩柱相对水深影响系数n2
        self.n2_dict = {"H/D": [[1,2,4,6,8,10,12,14], [0.76,0.78,0.82,0.85,0.89,0.93,0.97,1.00]]}
        # 墩柱水流力横向影响系数m2 = 1.0
        self.m2 = 1.0
        # 墩柱斜向水流作用影响系数m3 = 1.0
        self.m3 = 1.0

    # 计算钢管桩受水流影响长度
    def flow_press_L(self, HD1, HD2):
        # print(self.Hdelta, HD1, HD2)
        # 水比桩顶高
        if self.Hdelta > HD1:
            H = abs(HD1 - HD2)
        # 水在桩顶桩底之间
        elif self.Hdelta < HD1 and self.Hdelta > HD2:
            H = self.Hdelta - HD2
        # 水在桩底以下
        elif self.Hdelta < HD2:
            H = 0
        return H
    
    # 计算m1，S是前后墩之间的外边缘间距
    def cal_m1(self, L, D):
        k = L / D
        lst1 = self.m1_dict["L/D"][0]
        lst2 = self.m1_dict["L/D"][1]
        m1 = double_param_table_index(lst1, lst2, k, 2) 
        return m1
    
    # 计算n2，H是水深，也就是钢管桩受水流影响长度
    def cal_n2(self, H, D):
        k = H / D
        lst1 = self.n2_dict["H/D"][0]
        lst2 = self.n2_dict["H/D"][1]
        n2 = double_param_table_index(lst1, lst2, k, 2)
        return n2
    
    # 计算水流力
    def cal_Fw(self, D):
        # 前墩Cw
        Cw = 0.73
        Fw = Cw * 0.5 * self.ENV * self.V * self.V * 1 * D
        return round(Fw, 2)
    

def cal_ShuiLiuLi_JTS_144_1_2010(ENV, V, Hdelta, SP_H1, SP_H2, SP_D, SP_L, n):
    # 类参数
    classA = Flow_JTS_144_1_2010(ENV, V, Hdelta)
    # 外部参数
    # 外部参数——桩顶参数，单位m
    HD1 = float(SP_H1)
    # 外部参数——桩底参数，单位m
    HD2 = float(SP_H2)
    # 外部参数——钢管桩直径，单位m
    D = float(SP_D)
    # 外部参数——前后墩之间的外边缘间距L，单位m
    L = float(SP_L)
    # 计算钢管桩受水流影响长度
    H = classA.flow_press_L(HD1, HD2)
    # print("钢管桩受水流影响长度为:", round(H, 2))
    # 计算m1
    # 对于第一排桩，m1=1
    if SP_L == 0:
        m1 = 1.0
    else:
        m1 = classA.cal_m1(L, D)
    eq_m1 = pow(classA.cal_m1(L, D), n-1)
    # print("遮流影响系数m1为:", m1)
    # 计算n2
    n2 = classA.cal_n2(H, D)
    # print("墩柱相对水深影响系数n2为:", n2)
    # 计算水流力
    Fw = classA.cal_Fw(D)
    # 计算系数
    m2, m3, n1 = [1.0, 1.0, 1.0]
    Fw_lineload = round(eq_m1 * m2 * m3 * n1 * n2 * Fw, 2)
    # print("水流力Fw为(kN):", Fw)
    # lineload = ShuiLiuLi_Line_Loads(Fw, H)
    return Fw, Fw_lineload


class Wind_GB_55001_2021:
    # 基本参数有基本风速v0，地表分类DBFL，地形修正系数k，高度GD_Z，假定迎风贝雷排数n
    def __init__(self, v0, k, GD_Z, n, DBFL):
        self.v0 = float(v0)
        self.DBFL = DBFL
        self.k = float(k)
        self.GD_Z = float(GD_Z)
        self.n = n
        self.uz_dict = {"A:海面、海岸、开阔水面":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450,500,550],[1.09,1.28,1.42,1.52,1.67,1.79,1.89,1.97,2.05,2.12,2.18,2.23,2.46,2.64,2.78,2.91,2.91,2.91,2.91,2.91,2.91]],
                        "B:田野、乡村、丛林、平坦开阔地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450,500,550],[1.00,1.00,1.13,1.23,1.39,1.52,1.62,1.71,1.79,1.87,1.93,2.00,2.25,2.46,2.63,2.77,2.91,2.91,2.91,2.91,2.91]],
                        "C:树木及地层建筑密集区、平缓丘陵地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450,500,550],[0.65,0.65,0.65,0.74,0.88,1.00,1.10,1.20,1.28,1.36,1.43,1.50,1.79,2.03,2.24,2.43,2.60,2.76,2.91,2.91,2.91]],
                        "D:中高层建筑密集区、起伏较大的丘陵地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450,500,550],[0.51,0.51,0.51,0.51,0.51,0.60,0.69,0.77,0.84,0.91,0.98,1.04,1.33,1.58,1.81,2.02,2.22,2.40,2.58,2.74,2.91]],
                    }

    # 4.6.2基本风压应根据基本风速值进行计算，且其取值不得低于0.30kN/m2。
    def cal_w0(self):
        w0 = 0.5 * 1.25 * self.v0 * self.v0 / 1000
        if w0 < 0.3:
            w0 = 0.3
        return w0

    # 4.6.3风压高度变化系数应根据建设地点的地面粗糙度确定。标准地面粗糙度条件应为周边无遮挡的空旷平坦地形，其10m高处的风压高度变化系数应取1.0
    def cal_uz(self):
        if self.GD_Z <= 10.0:
            uz = 1.0
        else:
            uz = pow(self.GD_Z/10, 0.30)
        return uz

    # 按GB5009-2012的表8.2.1来确定uz
    def table_uz(self):
        # 高度表
        hlst = self.uz_dict[self.DBFL][0]
        # 地表分类下的uz表
        uzlst = self.uz_dict[self.DBFL][1]
        # 高度
        Z = self.GD_Z
        # 地表分类下对应高度的uz
        uz = double_param_table_index(hlst, uzlst, Z, 2)
        return uz 

    # 4.6.4体形系数，分为桁架的体形系数和钢管桩的体形系数。
    # 贝雷梁的体形系数，n为桁架的梠数
    def Truss_us(self):
        # 单梠桁架的体形系数，us按型钢杆件取，fai为单片贝雷梁的实面积比
        us = 1.3
        fai = 0.3
        ust = fai * us
        # n梠桁架的体形系数，η为查表的系数，b/h = 900/1500 < 1，取0.66
        η = 0.66
        ustw = ust * ((1-pow(η, self.n)) / (1-η))
        return ustw

    # 钢管桩的体形系数
    def Sp_us(self):
        us = 0.6
        return us
    
    # 4.6.5风荷载放大系数:主要受力结构的风荷载放大系数应根据地形特征、脉动风特性、结构周期、阻尼比等因素确定，其值不应小于1.2

    # 4.6.6地形修正系数:对于山峰和山坡等地形，其值不应小于1.0；对于山间盆地谷地，不应小于0.75；对于与风向一致的谷口、山口，不应小于1.20；其他情况，应取1.0。

    # 4.6.7风向影响系数：当有15年以上符合观测要求且可靠的风气象资料时，应按照极值理论的统计方法计算不同风向的风向影响系数。所有风向影响系数的最大值不应小于1.0，最小值不应小于0.8。其他情况，应取1.0

    # 计算风压
    def FengHeZai(self, us, uz, w0):
        Pa = 1.2 * 1.0 * self.k * us *uz * w0
        return round(Pa, 2)
    
def cal_FengYa_GB_55001_2021(v0, k, GD_Z, n, DBFL):
    # 基本信息
    ClassA = Wind_GB_55001_2021(v0, k, GD_Z, n, DBFL)
    # 基本风压
    w0 = ClassA.cal_w0()
    # print("基本风压w0为:", round(w0, 2))
    # 风压高度变化系数
    uz = ClassA.table_uz()
    # print("风压高度变化系数uz为:", uz)
    # 贝雷梁体形系数，n为贝雷梁梠数
    us1 = ClassA.Truss_us()
    # print("贝雷梁体形系数us为:", round(us1, 2))
    # 钢管桩体形系数
    us2 = ClassA.Sp_us()
    # print("钢管桩体形系数为:", round(us2, 2))
    # 地形修正系数k
    # 风向影响系数 = 1.0
    # 风荷载脉动增大效应 = 1.2
    # print("地形修正系数为:", k)
    # print("风向影响系数取1.0")
    # print("风荷载脉动增大效应取1.2")
    # 计算风压
    pa1 = ClassA.FengHeZai(us1, uz, w0)
    # print("贝雷梁上风压", round(pa1, 2))
    pa2 = ClassA.FengHeZai(us2, uz, w0)
    # print("钢管桩上风压", round(pa2, 2))

    return pa1, pa1, pa2

class Wind_JTG_T_3360_01_2018:

    # 基本参数有基本风速U10，地表分类DBFL，地形条件系数kt，高度GD_Z，水平加载长度L，假定迎风贝雷排数n，设计基准风速计算公式
    def __init__(self, U10, kt, Z, L, n, DBFL):
        # 基本风速，浮点
        self.U10 = float(U10)
        # 地表分类，字符串
        self.DBFL = DBFL
        # 地形条件系数kt
        self.kt = float(kt)
        # 由高度计算基准高度
        # self.JZGD_Z = float(GD_Z) * (2/3)
        self.JZGD_Z = float(Z)
        # 水平加载长度L
        self.L = float(L)
        # 贝雷梁迎风片数
        self.n = float(n)
        # 设计基准风速计算公式
        # self.formula = formula
        # 和地表分类有关的系数表，地表粗糙度系数a0和地表类别转换系数kc
        self.DBFL_dict = {"A:海面、海岸、开阔水面": [0.12, 1.174], "B:田野、乡村、丛林、平坦开阔地": [0.16, 1.0], "C:树木及地层建筑密集区、平缓丘陵地": [0.22, 0.785], "D:中高层建筑密集区、起伏较大的丘陵地": [0.30, 0.564]}
        # 和风险区域有关的系数表，抗风风险系数kf
        self.R_dict = {"R1": 1.05, "R2": 1.02, "R3": 1.0}
        # 地表类别转换及风速高度修正系数kh的参数表
        self.kh_dict = {"A:海面、海岸、开阔水面":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450],[1.08,1.17,1.23,1.28,1.34,1.39,1.42,1.46,1.48,1.51,1.53,1.55,1.62,1.68,1.73,1.77,1.77,1.77,1.77]],
                        "B:田野、乡村、丛林、平坦开阔地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450],[1.00,1.00,1.07,1.12,1.19,1.25,1.29,1.33,1.36,1.40,1.42,1.45,1.54,1.62,1.67,1.72,1.77,1.77,1.77]],
                        "C:树木及地层建筑密集区、平缓丘陵地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450],[0.86,0.86,0.86,0.92,1.00,1.06,1.12,1.16,1.20,1.24,1.27,1.30,1.42,1.52,1.59,1.66,1.71,1.77,1.77]],
                        "D:中高层建筑密集区、起伏较大的丘陵地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450],[0.79,0.79,0.79,0.79,0.85,0.85,0.91,0.96,1.01,1.05,1.09,1.13,1.27,1.39,1.48,1.57,1.64,1.71,1.77]],
                    }
        # 等效静阵风系数的参数表（主梁）
        self.Gv1_dict = {"A:海面、海岸、开阔水面":[[20,60,100,200,300,400,500,650,800,1000,1200,1500,2000],[1.29,1.28,1.26,1.24,1.23,1.22,1.21,1.20,1.19,1.18,1.17,1.16,1.15]],
                         "B:田野、乡村、丛林、平坦开阔地":[[20,60,100,200,300,400,500,650,800,1000,1200,1500,2000],[1.35,1.33,1.31,1.29,1.27,1.26,1.25,1.24,1.23,1.22,1.21,1.20,1.18]],
                         "C:树木及地层建筑密集区、平缓丘陵地":[[20,60,100,200,300,400,500,650,800,1000,1200,1500,2000],[1.49,1.48,1.45,1.41,1.39,1.37,1.36,1.34,1.33,1.31,1.30,1.29,1.26]],
                         "D:中高层建筑密集区、起伏较大的丘陵地":[[20,60,100,200,300,400,500,650,800,1000,1200,1500,2000],[1.56,1.54,1.51,1.47,1.44,1.42,1.41,1.39,1.37,1.35,1.34,1.32,1.30]],
                    }
        # 等效静阵风系数的参数表（墩柱）
        self.Gv2_dict = {"A:海面、海岸、开阔水面":[[40,60,80,100,150,200,300,400],[1.19,1.18,1.17,1.16,1.14,1.13,1.12,1.11]],
                         "B:田野、乡村、丛林、平坦开阔地":[[40,60,80,100,150,200,300,400],[1.24,1.22,1.20,1.19,1.17,1.16,1.14,1.13]],
                         "C:树木及地层建筑密集区、平缓丘陵地":[[40,60,80,100,150,200,300,400],[1.33,1.29,1.27,1.26,1.23,1.21,1.18,1.16]],
                         "D:中高层建筑密集区、起伏较大的丘陵地":[[40,60,80,100,150,200,300,400],[1.48,1.42,1.39,1.36,1.31,1.28,1.24,1.22]],
                    }

    # 基本风速U10——风险区域FXQY
    def FengXianQuYu(self):
        if self.U10 >= 32.6:
            FXQY = "R1"
        elif self.U10 < 32.6 and self.U10 >= 24.5:
            FXQY = "R2"
        elif self.U10 < 24.5:
            FXQY = "R3"
        return FXQY
    
    # 地表分类DBFL——地表粗糙度系数a0
    def CuCaoXiShu(self):
        a0 = self.DBFL_dict[self.DBFL][0]
        return a0
    
    # 地表分类DBFL——地表类别转换系数kc
    def DiBiaoLeiBieZhuanHuanXiShu(self):
        kc = self.DBFL_dict[self.DBFL][1]
        return kc

    # 风险区域FXQY——抗风风险系数kf
    def KangFengFengXianXiShu(self, R):
        kf = self.R_dict[R]
        return kf

    # 地表类别转换及风速高度修正系数kh——公式计算
    def cal_kh(self):
        a0 = self.DBFL_dict[self.DBFL][0]
        kc = self.DBFL_dict[self.DBFL][1]
        kh = kc * pow((self.JZGD_Z/10), a0)
        return round(kh, 2)

    # 地表类别转换及风速高度修正系数kh——表查询
    def table_kh(self):
        # 高度表
        hlst = self.kh_dict[self.DBFL][0]
        # kh值
        khlst = self.kh_dict[self.DBFL][1]
        # 基准高度
        Z = self.JZGD_Z
        # 检查基准高度是否为hlst中的值，是的话返回index并且在khlst中找到对应的kh值，不是的话线性内插
        kh = double_param_table_index(hlst, khlst, Z, 2)
        return kh

    # 等效静阵风系数Gv表查询
    def DengXiaoJingZhenFengXiShu(self):
        # 水平加载长度表
        llst = self.Gv1_dict[self.DBFL][0]
        # 主梁Gv1值表
        Gv1lst = self.Gv1_dict[self.DBFL][1]
        # 检查水平加载长度是否为llst中的值，是的话返回index并且在Gvlst中找到对应的Gv值，不是的话线性内插
        L = self.L
        # 对应地表类型对应水平长度的Gv1
        Gv1 = double_param_table_index(llst, Gv1lst, L, 3)
        # 桥墩高度表
        zlst = self.Gv2_dict[self.DBFL][0]
        # 桥墩Gv2值表
        Gv2lst = self.Gv2_dict[self.DBFL][1]
        # 高度用基准高度来算
        Z = self.JZGD_Z * 3/2
        # 对应地表类型对应桥墩高度的Gv2
        Gv2 = double_param_table_index(zlst, Gv2lst, Z, 3)
        return Gv1, Gv2

    # 计算Us10
    def cal_Us10(self, kc):
        return round(kc * self.U10, 1)

    # 计算基准高度Z处的设计基准风速，param控制计算公式的选取
    def cal_Ud(self, kf, kh, a0, Us10, formula):
        kt = self.kt
        U10 = self.U10
        # 基准高度Z处的设计基准风速
        if formula == "Ud = kf·(Z/10)^α0·Us10":
            Ud = kf * pow((self.JZGD_Z/10), a0) * Us10
        elif formula == "Ud = kf·kt·kh·U10":
            Ud = kf * kt * kh * U10
        return round(Ud, 1)

    # 计算等效静阵风风速Ug
    def cal_Ug(self, Gv, Ud):
        Ug = Gv * Ud
        return round(Ug, 1)

    # 计算主梁风荷载的风压大小Fg/D，D是主梁特征高度
    def ZhuLiang_FengHeZai(self, Ug, CH):
        # 压强
        Pa = 0.5 * 1.25 * pow(Ug, 2) * CH / 1000
        return round(Pa, 2)

    # 计算钢管桩风荷载的风压大小Fg/D，D是钢管桩直径
    def DunZhu_FengHeZai(self, Ug, CD):
        # 压强
        Pa = 0.5 * 1.25 * pow(Ug, 2) * CD / 1000
        return round(Pa, 2)
    
    # 贝雷迎风面积
    def Bailey_A(self):
        A = ((1-pow(0.66, self.n)) / (1-0.66)) * 0.3 * 4.5
        return round(A, 1)
    
    # 单片贝雷风荷载
    def Bailey_FF(self, Pa, A):
        return round(Pa*A, 2)

# 计算风压
def cal_FengYa_JTG_T_3360_01_2018(U10, kt, Z, L, n, CH, dbfl, formula):
    # 输入基本风速U10，地表分类DBFL，地形条件系数kt，高度GD_Z，水平加载长度L，假定迎风贝雷排数n，设计基准风速计算公式
    ClassA = Wind_JTG_T_3360_01_2018(U10, kt, Z, L, n, dbfl)
    # 风险区域
    R = ClassA.FengXianQuYu()
    # print("风险区域R为:", R)
    # 地面粗糙度系数
    a0 = ClassA.CuCaoXiShu()
    # print("地面粗糙度a0系数为:", a0)
    # 地表类别转换系数
    kc = ClassA.DiBiaoLeiBieZhuanHuanXiShu()
    # print("地表类别转换系数kc为:", kc)
    # 抗风风险系数
    kf = ClassA.KangFengFengXianXiShu(R)
    # print("抗风风险系数kf为:", kf)
    # 地表类别转换及风速高度修正系数
    kh = ClassA.cal_kh()
    # print("地形条件系数kh为:", kh)
    # kh = 2.0
    # 判断计算的地表类别转换及风速高度修正系数满不满足要求，若不满足则查表
    if kh < 1.0 or kh > 1.77:
        # print("kh不满足kh∈[1.0, 1.77]要求")
        kh = ClassA.table_kh()
        # print(f"查表得kh={kh}")
    # else:
    #     print("kh满足kh∈[1.0, 1.77]要求")
    # 等效静阵风系数
    Gv = ClassA.DengXiaoJingZhenFengXiShu()
    Gv1 = Gv[0]
    Gv2 = Gv[1]
    # print("主梁等效静阵风系数Gv:", Gv1)
    # print("桥墩等效静阵风系数Gv:", Gv2)
    # 设计基准风速和等效静阵风
    # 选择设计基准风速的计算公式
    Us10 = ClassA.cal_Us10(kc)
    Ud = ClassA.cal_Ud(kf, kh, a0, Us10, formula)
    Ug1 = ClassA.cal_Ug(Gv1, Ud)
    Ug2 = ClassA.cal_Ug(Gv2, Ud)
    # print("设计基准风速Ud:", Ud)
    # print("主梁等效静阵风风速Ug:", Ug1)
    # print("桥墩等效静阵风风速Ug:", Ug2)
    # 求风荷载
    # 主梁横向力系数
    ZL_CH = float(CH)
    # 贝雷梁横向力系数
    BL_CH = 1.7
    # 墩柱横向力系数
    DZ_CD = 0.6
    # 主梁风荷载
    Pa1 = ClassA.ZhuLiang_FengHeZai(Ug1, ZL_CH)
    # 贝雷风荷载
    Pa2 = ClassA.ZhuLiang_FengHeZai(Ug1, BL_CH)
    # 墩柱风荷载
    Pa3 = ClassA.DunZhu_FengHeZai(Ug2, DZ_CD)
    # print("主梁风压(kpa):", Pa1) 
    # print("墩柱风压(kpa):", Pa2)
    # 贝雷迎风面积计算
    A = ClassA.Bailey_A()
    # 单片贝雷风荷载
    FF = ClassA.Bailey_FF(Pa2, A)
    # return Pa1, Pa2
    # print(Gv1, Gv2, kf, kh, kc, Us10, Ud, a0, Ug1, Ug2, Pa1, Pa2, Pa3, A, FF)
    return Gv1, Gv2, kf, kh, kc, Us10, Ud, a0, Ug1, Ug2, Pa1, Pa2, Pa3, A, FF

class Wind_JTS_144_1_2010:

    # 基本参数有平均最大风速V，地表分类DBFL，基本风压增大或降低系数k，高度GD_Z，假定迎风贝雷排数n
    def __init__(self, V, k, GD_Z, n, DBFL):
        self.V = float(V)
        self.DBFL = DBFL
        self.k = float(k)
        self.GD_Z = float(GD_Z)
        self.n = n
        self.uz_dict = {"A:海面、海岸、开阔水面":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450],[1.17,1.38,1.52,1.63,1.80,1.92,2.03,2.12,2.20,2.27,2.34,2.40,2.64,2.83,2.99,3.12,3.12,3.12,3.12]],
                        "B:田野、乡村、丛林、平坦开阔地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450],[1.00,1.00,1.14,1.25,1.42,1.56,1.67,1.77,1.86,1.95,2.02,2.09,2.38,2.61,2.80,2.97,3.12,3.12,3.12]],
                        "C:树木及地层建筑密集区、平缓丘陵地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450],[0.74,0.74,0.74,0.84,1.00,1.13,1.25,1.35,1.45,1.54,1.62,1.70,2.03,2.30,2.54,2.75,2.94,3.12,3.12]],
                        "D:中高层建筑密集区、起伏较大的丘陵地":[[5,10,15,20,30,40,50,60,70,80,90,100,150,200,250,300,350,400,450],[0.62,0.62,0.62,0.62,0.62,0.73,0.84,0.93,1.02,1.11,1.19,1.27,1.61,1.92,2.19,2.45,2.68,2.91,3.12]],
                    }
        
    # 计算基本风压W0，不得小于0.3kpa
    def cal_W0(self):
        W0 = self.V * self.V / 1600
        if W0 < 0.3:
            W0 = 0.3
        # 返回值乘以增大或降低系数
        return W0*self.k
    
    # 体形系数，分为桁架的体形系数和钢管桩的体形系数。
    # 贝雷梁的体形系数，n为桁架的梠数
    def Truss_us(self):
        # 单梠桁架的体形系数，us按型钢杆件取，fai为单片贝雷梁的实面积比
        us = 1.3
        fai = 0.3
        ust = fai * us
        # n梠桁架的体形系数，η为查表的系数，b/h = 900/1500 < 1，取0.66
        η = 0.66
        ustw = ust * ((1-pow(η, self.n)) / (1-η))
        return ustw

    # 钢管桩的体形系数
    def Sp_us(self):
        us = 0.6
        return us
    
    # 表11.0.9
    def table_uz(self):
        # 高度表
        hlst = self.uz_dict[self.DBFL][0]
        # 地表分类下的uz表
        uzlst = self.uz_dict[self.DBFL][1]
        # 高度
        Z = self.GD_Z
        # 地表分类下对应高度的uz
        uz = double_param_table_index(hlst, uzlst, Z, 2)
        return uz
    
    # 计算风压
    def FengHeZai(self, us, uz, W0):
        Pa = us *uz * W0
        return round(Pa, 2)
    
def cal_FengYa_JTS_144_1_2010(V, k, GD_Z, n, DBFL):
    # 基本信息
    ClassA = Wind_JTS_144_1_2010(V, k, GD_Z, n, DBFL)
    # 基本风压
    W0 = ClassA.cal_W0()
    # print("基本风压w0为:", round(W0, 2))
    # 风压高度变化系数
    uz = ClassA.table_uz()
    # print("风压高度变化系数uz为:", uz)
    # 贝雷梁体形系数，n为贝雷梁梠数
    us1 = ClassA.Truss_us()
    # print("贝雷梁体形系数us为:", round(us1, 2))
    # 钢管桩体形系数
    us2 = ClassA.Sp_us()
    # print("钢管桩体形系数为:", round(us2, 2))
    # 计算风压
    pa1 = ClassA.FengHeZai(us1, uz, W0)
    # print("贝雷梁上风压", round(pa1, 2))
    pa2 = ClassA.FengHeZai(us2, uz, W0)
    # print("钢管桩上风压", round(pa2, 2))

    return pa1, pa1, pa2


# 钢结构设计规范——稳定系数
#flag 截面类型，可以输入的值是英文小写字符 "a"、"b"、"c"、"d"，或'a'这种写法
#length 计算长度，参见规范
#i 截面回转半径
#fy 材料牌号，q235钢材填235数字
#e 材料弹性 钢一般取2.06e+5
def stability_coefficient(flag, length, i, fy, e):
    lambda1 = length / i
    lambdan = lambda1 / math.pi * math.sqrt(fy / e)
    if flag == "a":
        alpha1 = 0.41
        alpha2 = 0.986
        alpha3 = 0.152
    elif flag == "b":
        alpha1 = 0.65
        alpha2 = 0.965
        alpha3 = 0.3
    elif flag == "c":
        alpha1 = 0.73
        if lambdan <= 1.05:
            alpha2 = 0.906
            alpha3 = 0.595
        else:
            alpha2 = 1.216
            alpha3 = 0.302
    elif flag == "d":
        alpha1 = 1.35
        if lambdan <= 1.05:
            alpha2 = 0.868
            alpha3 = 0.915
        else:
            alpha2 = 1.375
            alpha3 = 0.432
    if lambdan <= 0.215:
        fai = 1 - alpha1 * lambdan * lambdan
    else:
        tmp = alpha2 + alpha3 * lambdan + lambdan * lambdan
        fai = (tmp - math.sqrt(math.pow(tmp,2) - 4.0 * lambdan * lambdan)) / 2.0 / lambdan / lambdan
    return fai


# 钢管桩的计算参数
def SteelPile_param(model_post_dict):
    # 从字典中获取参数
    ele, Fx, My, Mz, D, t, L = model_post_dict["model_post_gangguanzhuang"]
    # 单位一定是KN和M, 见函数 Value_BeamForce 中的单位定义
    Fx = abs(round(float(Fx),1)) # Fx
    My = abs(round(float(My),1)) # My
    Mz = abs(round(float(Mz),1)) # Mz
    # 单位一定是N和MM, 见2.1小节单位修改函数
    D = round(float(D),1) # 直径
    t = round(float(t),1) # 壁厚
    L = round(float(L),1) # 桩长
    # 根据参数计算
    Area = math.pi*(D*D-(D-2*t)*(D-2*t))/4
    Ix = math.pi*(D**4-(D-2*t)**4)/64
    ix = math.sqrt(Ix/Area)
    Phi = stability_coefficient("b", L, ix, 235, 206000)
    Mm = math.sqrt(math.pow(My, 2) + math.pow(Mz, 2))
    Wx = Ix/(D/2)
    quote_lambda = L/ix
    Nex = (math.pi**2)*206000*Area/(1.1*(quote_lambda**2))
    sigma = ((Fx*1000/(Phi*Area))+((1*Mm*1000000)/(1.15*Wx*(1-0.8*(Fx*1000/Nex)))))
    return [Fx, Mm, L, Area, Ix, ix, Phi, Wx, quote_lambda, Nex, sigma]


# 基础及其上土体自重、基础偏心、基础地基应力、软弱下卧层
# 基础及其上土的自重, h1是基础底标高, 用于计算土的厚度
def foundation_Gk(x, y, z, h1, gc, gs):
    # 单位kN m
    Gc = round(x * y * z * gc, 1)
    # 土厚
    hs0 = round(abs(h1) - z, 1)
    if hs0 > 0:
        hs = hs0
    else:
        hs = 0
    Gs = round(x * y * hs * gs, 1)
    Gk = Gc + Gs
    return round(Gk, 1)


# 轴心受压pk, xyz是基础尺寸
def foundation_pk(x, y, Gk, Fk):
    Ak = x * y
    pk = (Fk + Gk) / Ak
    return [round(Ak, 1), round(pk, 1)]


# 偏心计算，基础抗弯方向长度b，垂直于力矩作用方向的长度为l，当My / (Fk+Gk) > b/6时，基础大偏心，pk在基底面上影响的范围会发生变化
def cal_Eccentricity(Gk, Fk, M, x, y, param):
    e = M/(Gk+Fk)
    if param == "Mx":
        a = x/2 - e
        if e > x/6:
            return [round(e,2), round(a,2), "大偏心"]
        else:
            return [round(e,2), round(a,2), "小偏心"]
    if param == "My":
        a = y/2 - e
        if e > y/6:
            return [round(e,2), round(a,2), "大偏心"]
        else:
            return [round(e,2), round(a,2), "小偏心"]
        

# 偏心受压pkmax和pkmin, xyz是基础尺寸
def foundation_pkmax_pkmin(x, y, Gk, Fk, Mk, param):
    Ak = x * y
    if param == "Mx":
        Wk = (y * x * x) / 6
    if param == "My":
        Wk = (x * y * y) / 6
    pkmax = (Fk + Gk) / Ak + Mk / Wk
    pkmin = (Fk + Gk) / Ak - Mk / Wk
    return [round(Wk, 3), round(pkmax, 1), round(pkmin, 1)]


# 大偏心时基底应力计算
def foundation_pkmax_Large_Eccentricity(Gk, Fk, x, y, a, param):
    if param == "Mx":
        pkmax = 2*(Fk+Gk) / (3*y*a)
        return round(pkmax, 1)
    if param == "My":
        pkmax = 2*(Fk+Gk) / (3*x*a)
        return round(pkmax, 1)
    
    
# 有软弱下卧层pcz, h为软弱下卧层顶标高, g为土容重
def foundation_pcz(h, g):
    pcz = abs(g * h)
    return round(pcz, 1)


# 有软弱下卧层pz, xyz是基础平面尺寸, gc是基础容重, h1为基础底标高, h2为软弱下卧层顶标高, gs为土容重, angle为扩散角
def foundation_pz(x, y, z, pk, h1, h2, gs, angle):
    # 土厚
    hs0 = round(abs(h1) - z, 1)
    if hs0 > 0:
        hs = hs0
    else:
        hs = 0
    if hs != 0:
        pc = abs(gs * hs)
    else:
        pc = 0
    delta_h = abs(h2 - h1)
    tan_angle = round(math.tan(math.radians(angle)), 1)
    Az = (x + 2*delta_h*tan_angle) * (y + 2*delta_h*tan_angle)
    pz = x * y * (pk - pc) / Az
    return [round(pc, 1), round(pz, 1)]


# 计算基底净反力
def cal_p(ksi, pk, l, b, Gk):
    p = ksi * (pk - Gk/(l*b))
    return round(p, 1)


def cal_pmax(ksi, pkmax, l, b, Gk):
    pmax = ksi * (pkmax - Gk/(l*b))
    return round(pmax, 1)


def cal_pmin(ksi, pkmin, l, b, Gk):
    pmin = ksi * (pkmin - Gk/(l*b))
    if pmin < 0:
        pmin = 0
    return round(pmin, 1)


# 基础抗剪验算
# 判断是否需要计算抗剪，l和b是基础的两边长，h为基础高度，as0为保护层厚度，at是对应基础短边的柱边, bt是对应基础长边的柱边, 对于圆柱转换成0.8d边长的正方形柱
def if_Vs(l, b, h, as0, at, bt, D):
    if at == 0 and bt == 0 and D != 0:
        at = 0.8*D
        bt = 0.8*D
    else:
        return None
    if min([l,b]) <= at + 2*(h-as0/1000):
        return True
    else:
        return False
    

# 计算受剪切承载力截面高度影响系数
def Vs_Bhs(h0):
    if h0 < 800:
        h = 800
    elif h0 > 2000:
        h = 2000
    else:
        h = h0
    Bhs = math.pow((800/h), 0.25)
    return round(Bhs, 2)


# 计算荷载的基本组合作用下，基础受到的剪力Vs，b是力矩方向上的基础边长，l是垂直力矩方向上的基础边长，at是对应基础短边的柱边, bt是对应基础长边的柱边, 对于圆柱转换成0.8d边长的正方形柱
def cal_Vs(pmax, l, b, at, bt):
    As1 = b * (l-at)/2 # 上下
    As2 = l * (b-bt)/2 # 左右
    Vs1 = pmax * As1 # 上下
    Vs2 = pmax * As2 # 左右
    return [round(As1, 3), round(As2, 3), round(Vs1, 1), round(Vs2, 1)]


# 计算基础受到的Vs，at是对应基础短边的柱边, bt是对应基础长边的柱边, 对于圆柱转换成0.8d边长的正方形柱
def foundation_Vs(x, y, at, bt, D, pmax):
    if at == 0 and bt == 0 and D != 0:
        at = 0.8*D
        bt = 0.8*D
    Vs = cal_Vs(pmax, x, y, at, bt)
    return Vs


# 计算容许值0.7*Bhs*ft*A0
def foundation_Allowable_Vs(l, b, h, a0, ft):
    # 截面有效高度
    h0 = h*1000 - a0
    # 截面高度影响系数
    Bhs = Vs_Bhs(h0)
    # 受剪面积
    A1 = b*1000 * h0 # 上下
    A2 = l*1000 * h0 # 左右
    # 公式计算值
    Allowable_Vs_value1 = 0.7 * Bhs * ft * A1 / 1000 # 上下
    Allowable_Vs_value2 = 0.7 * Bhs * ft * A2 / 1000 # 左右
    return [round(Bhs,1), round(A1,1), round(A2,1), round(Allowable_Vs_value1, 1), round(Allowable_Vs_value2, 1)]


# 抗冲切计算
# 计算Bhp
def Fl_Bhp(h):
    if h <= 800:
        Bhp = 1
    elif h >= 2000:
        Bhp = 0.9
    else:
        Bhp = 1 - (h-800)/1200*0.1
    return round(Bhp, 3)


# 计算am, 抗冲切计算中取最不利边长，上下左右四个方向的抗冲切承载力均取短边对应的am
def cal_am(at, h0):
    ab = at + 2*h0
    am = (at + ab) / 2
    return am


# 计算Al, at是对应基础短边的柱边, bt是对应基础长边的柱边, 对于圆柱转换成0.8d边长的正方形柱
def cal_Al(at, bt, l, b, h ,a0):
    # 有效高度
    h0 = (h*1000 - a0)/1000
    # 柱扩散后短边边长
    sl = at + 2*h0
    # 柱扩散后长边边长
    sb = bt + 2*h0
    # l边留距
    xl = (l-at)/2
    # b边留距
    xb = (b-bt)/2
    # print(sl, sb, xl, xb)
    if xb >= xl:
        if sb >= b:
            x_Ab, x_Al, s_Ab, s_Al = [0, 0, 0, 0]
            Ab, Al = [0, 0]# 面积均为0
        else:
            if sl >= l:
                x_Ab, x_Al, s_Ab, s_Al = [0, l, sb, sl]
                Ab, Al = [0, l*(b-sb)/2]# Ab面积为0, Al面积为正方形
            else:
                x_Ab, x_Al, s_Ab, s_Al = [bt+l-at, l, sb, sl]
                print(x_Ab, x_Al, s_Ab, s_Al)
                Ab, Al = [(s_Ab+x_Ab)*(xl-h0)/2, (s_Al+x_Al)*(xl-h0)/2+(l*(xb-xl))]# Ab是梯形，Al是梯形+正方形
    if xb < xl:
        if sl >= l:
            x_Ab, x_Al, s_Ab, s_Al = [0, 0, 0, 0]
            Ab, Al = [0, 0]# 面积均为0
        else:
            if sb >= b:
                x_Ab, x_Al, s_Ab, s_Al = [b, 0, sb, sl]
                Ab, Al = [0, b*(l-sl)/2]# Ab面积为正方形, Al面积为0
            else:
                x_Ab, x_Al, s_Ab, s_Al = [b, at+b-bt, sb, sl]
                Ab, Al = [(s_Ab+x_Ab)*(xb-h0)/2+(b*(xl-xb)), (s_Al+x_Al)*(xb-h0)/2]# Ab是梯形+正方形，Al是梯形
    return [round(Ab, 3), round(Al, 3)]


# 计算Fl
def cal_Fl(pmax, Al):
    Fl = pmax * Al
    return Fl


# 计算基础受到的Fl, l是短边, b是长边
def foundation_Fl(pmax, at, bt, D, l, b, h, a0):
    if at == 0 and bt == 0 and D != 0:
        at = 0.8*D
        bt = 0.8*D
    # 冲切基底面积
    Al1, Al2 = cal_Al(at, bt, l, b, h ,a0)
    print("冲切基底面积: ", Al1, Al2)
    # 冲切破坏锥体最不利一侧计算长度
    Fl1 = cal_Fl(pmax, Al1)
    Fl2 = cal_Fl(pmax, Al2)
    return [round(Al1, 3), round(Al2, 3), round(Fl1, 1), round(Fl2, 1)]


# 计算容许值0.7*Bhp*ft*am*h0, at是对应基础短边的柱边, bt是对应基础长边的柱边, 对于圆柱at=bt
def foundation_Allowable_Fl(ft, at, bt, h, D, a0):
    if at == 0 and bt == 0 and D != 0:
        at = 0.8*D
        bt = 0.8*D
    # 有效高度
    h0 = h*1000 - a0
    # 受冲切承载力截面高度影响系数
    Bhp = Fl_Bhp(h)
    # 冲切破坏锥体最不利一侧计算长度
    am = cal_am(min(at*1000, bt*1000), h0)# 短边
    # 容许值
    Allowable_Fl = 0.7 * Bhp * ft * am * h0 / 1000
    return [round(h0, 1), round(Bhp, 1), round(am, 1), round(Allowable_Fl, 1)]


# 基础抗弯及配筋
# 短边方向M1
def cal_M(x, y, a1, aa, bb, p, pmax, pmin, type_M):
    if type_M == "Mx":
        M1 = 1/12 * a1**2 * ((2*y+aa) * (pmax+p) + y * (pmax-p))
        M2 = 1/48 * (y-aa)**2 * (2*x+bb) * (pmax + pmin)
    if type_M == "My":
        M1 = 1/12 * a1**2 * ((2*x+aa) * (pmax+p) + x * (pmax-p))
        M2 = 1/48 * (x-aa)**2 * (2*y+bb) * (pmax + pmin)
    return [round(M1, 1), round(M2, 1)]


# 计算抗弯公式中的几何参数
def cal_M_param(x, y, at, bt, D, type_M):
    if at == 0 and bt == 0 and D != 0:
        at = 0.8*D
        bt = 0.8*D
    if type_M == "Mx":
        # 悬臂长度
        a1 = (x-at) / 2
        # 扩散长度
        aa = bt
        bb = at
    if type_M == "My":
        # 悬臂长度
        a1 = (y-bt) / 2
        # 扩散长度
        aa = at
        bb = bt
    return round(a1, 3), round(aa, 3), round(bb, 3)


# 根据弯矩计算需要的配筋面积
def cal_As(M, fy, x, y, h, a0, type_M):
    h0 = h*1000 - a0
    if type_M == "Mx":
        As = (M*1000000 / (0.9*fy*h0)) / y
    if type_M == "My":
        As = (M*1000000 / (0.9*fy*h0)) / x
    print("弯矩对应每延米钢筋截面面积为: ", round(As, 1))
    return round(As, 1)


# 计算最小配筋率对应的钢筋截面面积
def min_p_As(h, a0):
    p = 0.0015
    h0 = h*1000 - a0
    As = 1000 * h0 * p
    print("最小配筋率下每延米钢筋截面面积为: ", round(As, 1))
    return round(As, 1)


# 计算范围内所有是2的倍数的整数
def double_int_in_range(float1, float2, param):
    amin = min(float1, float2)
    amax = max(float1, float2)
    if int(amin * 10) % 10 == 0:
        a1 = int(amin)
    else:
        a1 = int(amin) + 1
    a2 = int(amax)
    alst = [x for x in range(a1, a2 + 1) if x % param == 0]
    
    return alst


# 选取钢筋, l是基础短边, b是基础长边, s是钢筋间距, d是钢筋直径, As是需要的面积
def cal_rebar_arrangement(h, a0, s1, s2, d1, d2, As):
    h0 = h*1000 - a0 - 5
    smin = min(s1, s2)
    smax = max(s1, s2)
    dmin = min(d1, d2)
    dmax = max(d1, d2)
    # 最大间距不宜大于200mm
    if smax > 200:
        smax = 200
    if smax < 100:
        smax = 100
    # 最小间距不宜小于100mm
    if smin > 200:
        smin = 200
    if smin < 100:
        smin = 100
    # 最小直径不宜小于10mm
    if dmin < 10:
        dmin = 10
    # 可选间距
    s_lst = list(reversed(double_int_in_range(smin, smax, 10)))
    # 钢筋种类
    d_lst = double_int_in_range(dmin, dmax, 2)
    # 计算最小间距最大直径
    As_dmax_smin = ((math.pi * dmax**2) / 4) * (int(1000 / smin))
    # 用户给与的直径和间距默认可以满足要求
    found_As = True
    found_p = True
    # 最大也不满足
    if As_dmax_smin < As:
        print("最大面积不满足")
        found_As = False
    p_dmax_smin = (As_dmax_smin / (1000*h0))
    if p_dmax_smin*100 < 0.15:
        print("面积取到最大不满足最小配筋率")
        found_p = False
    # 其他情况
    else:
        # 从最小的直径开始，计算每一种间距对应的面积
        found = False
        for num_d in d_lst:
            for num_s in s_lst:
                As_select = ((math.pi * num_d**2) / 4) * (int(1000 / num_s))
                p_select = As_select/(1000*h0)
                if As_select >= As and p_select*100 >= 0.15:
                    s_select, d_select = [num_s, num_d]
                    found = True
                    break
            if found:
                break
    if found_As == False or found_p == False or found == False:
        return None
    else:
        return [s_select, d_select, p_select, As_select]
    

# 原函数 concrete_rebar_f
# 混凝土等级所对应的参数
def concrete_f(concrete_grade):
    concrete_value_dict = {
        'compressive': {'C15': 7.2,'C20': 9.6,'C25': 11.9,'C30': 14.3,'C35': 16.7,'C40': 19.1,'C45': 21.1,'C50': 23.1,'C55': 25.3,'C60': 27.5,'C65': 29.7,'C70': 31.8,'C75': 33.8,'C80': 35.9},
        'tensile': {'C15': 0.91,'C20': 1.10,'C25': 1.27,'C30': 1.43,'C35': 1.57,'C40': 1.71,'C45': 1.80,'C50': 1.89,'C55': 1.96,'C60': 2.04,'C65': 2.09,'C70': 2.14,'C75': 2.18,'C80': 2.22},
        'concrete_Ec':{'C15': 22000,'C20': 25500,'C25': 28000,'C30': 30000,'C35': 31500,'C40': 32500,'C45': 33500,'C50': 34500,'C55': 35500,'C60': 36000,'C65': 36500,'C70': 37000,'C75': 37500,'C80': 38000}
    }
    fc = concrete_value_dict['compressive'][concrete_grade]
    ft = concrete_value_dict['tensile'][concrete_grade]
    Ec = concrete_value_dict['concrete_Ec'][concrete_grade]
    return [fc, ft, Ec]


# 原函数 concrete_rebar_f
# 钢筋等级所对应的参数  
def rebar_f(rebar_grade):
    rebar_value_dict = {
        'rebar_fy':{'HPB300': 270,'HRB335':300,'HRB400':360,'HRBF400':360,'RRB400':360,'HRB500':435,"HRBF500":435},
        'rebar_Es':{'HPB300': 210000,'HRB335':200000,'HRB400':200000,'HRBF400':200000,'RRB400':200000,'HRB500':200000,"HRBF500":200000}
    }
    fy = rebar_value_dict['rebar_fy'][rebar_grade]
    Es = rebar_value_dict['rebar_Es'][rebar_grade]

    return [fy, Es]


# 单桩竖向承载力
# 侧阻效应系数
def Side_Resistance_Factor(s_type, d):
    Side_Resistance_Factor_Lst = []
    for x in s_type:
        if x == "黏土":
            Side_Resistance_Factor_Lst.append(min(round((0.8/d)**(1/5),2), 1.0))
        elif x == "砂土":
            Side_Resistance_Factor_Lst.append(min(round((0.8/d)**(1/3),2), 1.0))
    return Side_Resistance_Factor_Lst


# 端阻效应系数
def Base_Resistance_Factor(s_type, d):
    Base_Resistance_Factor_Lst = []
    for x in s_type:
        if x == "黏土":
            Base_Resistance_Factor_Lst.append(min(round((0.8/d)**(1/4),2), 1.0))
        elif x == "砂土":
            Base_Resistance_Factor_Lst.append(min(round((0.8/d)**(1/3),2), 1.0))
    return Base_Resistance_Factor_Lst


# 打入桩Qsk
def JGJ94_Qsk_DrivenPile(d, qsk_lst, l_lst):
    u = math.pi * d
    Qsk = u * sum(list(map(mul, qsk_lst, l_lst)))
    return round(Qsk, 1)


# 钢管桩Qpk
def JGJ94_Qpk_steelpile(d, qpk, hb):
    # 桩端土塞效应系数
    lambda_p = round(min(0.16*hb/d, 0.8),1)
    Ap = (math.pi * d**2) / 4
    Qpk = lambda_p * Ap * qpk
    return round(Qpk, 1)


# 预制桩Qpk
def JGJ94_Qpk_precastpile(d, t, qpk, hb):
    # 空心桩内径
    d1 = d-2*t
    # 空心装桩端净面积
    Aj = math.pi * (d**2 - d1**2) / 4
    # 空心桩敞口面积
    Apl = math.pi * d1**2 / 4
    # 桩端土塞效应系数
    lambda_p = round(min(0.16*hb/d1, 0.8),1)
    Qpk = qpk * (Aj + lambda_p * Apl)
    return round(Qpk, 1)


# 灌注桩Qsk
def JGJ94_Qsk_DrilledShaft(d, qsk_lst, l_lst, param_lst):
    u = math.pi * d
    Qsk = u * sum(list(map(mul, (map(mul, qsk_lst, l_lst)), param_lst)))
    return round(Qsk, 1)


# 灌注桩Qpk
def JGJ94_Qpk_DrilledShaft(d, qpk, param_lst):
    Ap = (math.pi * d**2) / 4
    Qpk = Ap * qpk * param_lst[-1]
    return round(Qpk, 1)


# Quk
def JGJ94_Quk(Qsk, Qpk):
    return round(Qsk+Qpk, 1)


# Ra
def JGJ94_Ra(Quk, K = 2):
    return round(Quk / K, 1)


# 配筋率小于0.65%的灌注桩的单桩水平承载力
# 桩身换算截面受拉边缘的截面模量, 单位mm3
def Section_Modulus_of_Transformed_Section_at_Tension_Edge_of_Pile_Shaft(d, Ec, Es, pg, d_as):
    W0 = math.pi * d * (d**2 + 2 * (Es/Ec - 1) * (pg/100) * (d-2*d_as)**2) / 32
    return round(W0, 1)


# 桩身换算截面积，单位mm2
def Transformed_Section_Area_of_Pile_Shaft(d, Ec, Es, pg):
    An = (math.pi * d**2 / 4) * (1 + (Es/Ec - 1) * (pg/100))
    return round(An, 1)


# Rha计算，单位kN
def JGJ94_Rha1(alpha, rm, ft, W0, vm, pg, N, An, param = 0):
    if param == 0: # 竖向力为压力
        Rha = (0.75 * (alpha/1000) * rm * ft * W0 / vm) * (1.25 + 22 * (pg/100)) * (1 + (0.5 * (N*1000))/(rm * ft * An)) / 1000
    elif param == 1: # 竖向力为拉力
        Rha = (0.75 * (alpha/1000) * rm * ft * W0 / vm) * (1.25 + 22 * (pg/100)) * (1 - (1.0 * (N*1000))/(rm * ft * An)) / 1000
    return round(Rha, 1) # 单位kN


# 钢桩、预制桩、配筋率大于等于0.65%的灌注桩的单桩水平承载力
# m计算, vb = 10 # 水平位移值，单位mm
def Subgrade_Reaction_Coefficient(friction, Cohesion, vb):
    m_lst = list(map(lambda x, y: (0.2*(x**2)-x+y)/vb*1000, friction, Cohesion))# 桩侧土水平抗力系数的比例系数m，单位KN/m4
    return m_lst


# α计算，桩的水平变形系数
def Displacement_Coefficient(m, b0, E, I):
    alpha = ((m*b0) / (E*I)) ** 0.2
    print(f"m: {m} (kN/m4), b0: {b0} (m), E: {E} (kN/m2), I: {I} (m4)")
    return round(alpha, 3) # 单位1/m


# vm及vx表
def JGJ94_Table_Vm_Vs(alpha, h, constraint = "铰接"):
    # 表值
    Vm_dict = {
        "铰接":{'4.0':0.768, '3.5':0.750, '3.0':0.703, '2.8':0.675, '2.6':0.639, '2.4':0.601},
        "固接":{'4.0':0.926, '3.5':0.934, '3.0':0.967, '2.8':0.990, '2.6':1.108, '2.4':1.045}
        }
    Vx_dict = {
        "铰接":{'4.0':2.441, '3.5':2.502, '3.0':2.727, '2.8':2.905, '2.6':3.163, '2.4':3.526},
        "固接":{'4.0':0.940, '3.5':0.970, '3.0':1.028, '2.8':1.055, '2.6':1.079, '2.4':1.095}
        }
    # 根据 alpha * h 的值和约束情况查表
    alpha_h = min(max(alpha * h, 2.4), 4.0)
    # 所有的key按从小到大排序
    Vm_keys = sorted(float(k) for k, v in Vm_dict[constraint].items())
    Vx_keys = sorted(float(k) for k, v in Vx_dict[constraint].items())
    # 判断范围
    Vm_left, Vm_right = max(x for x in Vm_keys if x <= alpha_h), min(x for x in Vm_keys if x >= alpha_h)
    Vx_left, Vx_right = max(x for x in Vx_keys if x <= alpha_h), min(x for x in Vx_keys if x >= alpha_h)
    # 内插
    if alpha_h in Vm_keys:
        Vm = Vm_dict[constraint][str(alpha_h)]
    else:
        Vm_a, Vm_b = Vm_dict[constraint][str(Vm_left)], Vm_dict[constraint][str(Vm_right)]
        Vm = Vm_a + (Vm_b-Vm_a) / (Vm_right-Vm_left) * (alpha_h-Vm_left)
        
    if alpha_h in Vx_keys:
        Vx = Vx_dict[constraint][str(alpha_h)]
    else:
        Vx_a, Vx_b = Vx_dict[constraint][str(Vx_left)], Vx_dict[constraint][str(Vx_right)]
        Vx = Vx_a + (Vx_b-Vx_a) / (Vx_right-Vx_left) * (alpha_h-Vx_left)
    return [round(Vm, 3), round(Vx, 3), round(alpha_h, 1)]


# Rha计算，单位kN
def JGJ94_Rha2(alpha, E, I, vx, vb):
    x0a = vb / 1000
    Rha = 0.75 * (alpha**3 * E * I) / vx * x0a
    return round(Rha, 1) # 单位kN

# ===========================================================================
# 以下为新增函数/类 260814
# ===========================================================================

# 建筑地基基础设计规范 GB50007-2011
GB50007_2011_5_2_2_1 = r'p_k=\frac{\left(F_k+G_k\right)}{A}'
GB50007_2011_Eccentricity = r'e=\frac{M_k}{\left(F_k+G_k\right)}'
GB50007_2011_5_2_2_2 = r'p_{kmax}=\frac{\left(F_k+G_k\right)}{A}+\frac{M_k}{W}'
GB50007_2011_5_2_2_3 = r'p_{kmin}=\frac{\left(F_k+G_k\right)}{A}-\frac{M_k}{W}'
GB50007_2011_5_2_2_4 = r'p_{kmax}=\frac{2\left(F_k+G_k\right)}{3la}'
GB50007_2011_5_2_7_3 = r'p_z=\frac{b\left(p_k-p_c\right)}{\left(b+2ztan\theta\right)\left(l+2ztan\theta\right)}'
GB50007_2011_5_2_8_pj = r'p_j=K_s\left(p_k-\frac{G_k}{A}\right)'
GB50007_2011_5_2_11_pjmax = r'p_{jmax}=K_s\left(p_{kmax}-\frac{G_k}{A}\right)'
GB50007_2011_5_2_11_pjmin = r'p_{jmin}=K_s\left(p_{kmin}-\frac{G_k}{A}\right)'


# 建筑桩基技术规范 JGJ94-2008
JGJ94_2008_5_3_7_1 = r'Q_{uk}=Q_{sk}+Q_{pk}=\mu\sum{q_{sik}l_i+\lambda_pq_{pk}A_p}'
JGJ94_2008_5_3_8_1 = r'Q_{uk}=Q_{sk}+Q_{pk}=\mu\sum{q_{sik}l_i+q_{pk}\left(A_j+\lambda_pA_{pl}\right)}'
JGJ94_2008_5_3_6 = r'Q_{uk}=Q_{sk}+Q_{pk}=\mu\sum{\psi_{si}q_{sik}l_i+\psi_pq_{pk}A_p}'
JGJ94_2008_5_2_2 = r'R_a=\frac{1}{K}Q_{uk}'
JGJ94_2008_5_7_2_1_W0 = r'W_0=\frac{\pi d}{32}\left[d^2+2\left(\alpha_E-1\right)\rho_gd_0^2\right]'
JGJ94_2008_5_7_2_1_An = r'A_n=\frac{\pi d^2}{4}\left[1+\left(\alpha_E-1\right)\rho_g\right]'
JGJ94_2008_5_7_2_1 = r'R_{ha}=\frac{0.75\alpha\gamma_mf_tW_0}{\nu_M}\left(1.25+22\rho_g\right)\left(1\pm\frac{\zeta_N\bullet N}{\gamma_mf_tA_n}\right)'
JGJ94_2008_5_7_5 = r'\alpha=\sqrt[5]{\frac{mb_0}{EI}}'
JGJ94_2008_5_7_2_2 = r'R_{ha}=0.75\frac{\alpha^3EI}{v_x}x_{0a}'

# 公路桥梁抗风设计规范 JTG/T 3360-01-2018
JTG_3360_01_2018_4_2_4 = r'U_{s10}=k_cU_{10}'
JTG_3360_01_2018_4_2_6_1 = r'U_d=k_f\left(\frac{Z}{10}\right)^{\alpha_0}U_{s10}'
JTG_3360_01_2018_4_2_6_2 = r'U_d=k_fk_tk_hU_{10}'
JTG_3360_01_2018_5_2_1 = r'U_g=G_VU_d'
JTG_3360_01_2018_5_3_1 = r'F_g/D=\frac{1}{2}\rho U_g^2C_H'

# 港口工程荷载规范 JTS 144-1-2010
JTS_144_1_2010_13_0_1 = r'F_\omega=C_\omega\frac{\rho}{2}V^2A'

# 建筑基坑支护技术规程 JGJ 120-2012
JGJ_120_2012_4_1_6 = r'm=\frac{0.2\varphi^2-\varphi+c}{v_b}'
JGJ_120_2012_4_2_2 = r'R_{ha}=0.75\frac{\alpha^3EI}{v_x}x_{0a}'
JGJ_120_2012_4_2_4 = r'R_{ha}=\frac{0.75\alpha\gamma_mf_tW_0}{\nu_M}\left(1.25+22\rho_g\right)'

# 钢结构设计标准 GB50017-2017
GB50017_2017_8_2_1_1 = r'\sigma=\frac{N}{\varphi A}+\frac{\beta M}{\gamma_mW(1-0.8\frac{N}{N_{Ex}^\prime})}'
GB50017_2017_8_2_1_2 = r'N_{EX}^\prime=\frac{\pi^2EA}{1.1\ \lambda^2}'
