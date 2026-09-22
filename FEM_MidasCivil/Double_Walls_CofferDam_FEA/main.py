# 2. 第三方库
import win32com.client 

# 3. 本地模块
from FEM_MidasCivil.Double_Walls_CofferDam_FEA.handle_data_to_mct import makeup_finallst

if __name__ == "__main__":
    acadapp = win32com.client.Dispatch("AutoCAD.Application")
    makeup_finallst(acadapp)
