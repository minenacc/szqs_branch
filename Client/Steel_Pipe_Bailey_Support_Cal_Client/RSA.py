# 生成申请码
import uuid
import socket
import platform
import random

# RSA加密
import rsa
import base64
import winreg

# 生成RSA密钥对
# def RSA():
#     (pubkey, privkey) = rsa.newkeys(512)
#     # 打印公钥和私钥
#     print("Public Key:", pubkey)
#     print("Private Key:", privkey)


# 将信息写入注册表
def write_Register(path, key, value):
    # 打开或创建注册表项
    registry_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
    # 设置键值
    winreg.SetValueEx(registry_key, key, 0, winreg.REG_SZ, value)
    # 关闭注册表项
    winreg.CloseKey(registry_key)

# 读取注册表中的信息
def read_Register(path, key_handle):
    try:
        # 打开或创建注册表项
        registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, path)
        # 设置键值
        value, value_name = winreg.QueryValueEx(registry_key, key_handle)
        # 关闭注册表项
        winreg.CloseKey(registry_key)
        # 返回值
        return value
    except FileNotFoundError:
        return "NotFound"

# 删除主键的值
def delete_registry_value(path, key_handle):
    try:
        # 打开或创建注册表项
        registry_key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, path)
        # 删除键值
        winreg.DeleteValue(registry_key, key_handle)
        # 关闭注册表项
        winreg.CloseKey(registry_key)
        return True
    except FileNotFoundError:
        return "NotFound"

# 私钥还原注册码
def Recover_code(registration_code, privkey, path, key_handle):
    # 解密注册码（仅供演示，实际应用中不应公开私钥）
    decrypted_info = rsa.decrypt(base64.b64decode(registration_code), privkey).decode()
    # print(decrypted_info)
    
    # 获取注册机中指定路径的信息
    # user_info = read_Register(path, key_handle)
    # print(user_info)
    
    # 判断解密后的数据和原数据是否一致
    # 给用户的数据既是registration_code
    # if decrypted_info == user_info:
    #     print("success")
    # else:
    #     print("defult")
    return decrypted_info


# # 写入注册表
# Reg_path = "Software\\bc_design_copilot"
# Reg_key = "AutoDoc"
# Reg_value = get_unique_id()
# write_Register(Reg_path, Reg_key, Reg_value)
# # 获取注册码
# zhucema = Register(pubkey, Reg_value)
# # 检查还原后的注册码是否与原信息一致
# Recover_code(zhucema, privkey, Reg_path, Reg_key)


# 获取注册码
# Reg_value = "LAPTOP-PD4B5JLV-Windows-11-10.0.22631"
# zhucema = Register(pubkey, Reg_value)

import wmi

class Hardware:
    @staticmethod
    def get_cpu_sn():
        """
        获取CPU序列号
        :return: CPU序列号
        """
        c = wmi.WMI()
        for cpu in c.Win32_Processor():
            # print(cpu.ProcessorId.strip())
            return cpu.ProcessorId.strip()

    @staticmethod
    def get_baseboard_sn():
        """
        获取主板序列号
        :return: 主板序列号
        """
        c = wmi.WMI()
        for board_id in c.Win32_BaseBoard():
            # print(board_id.SerialNumber)
            return board_id.SerialNumber

    @staticmethod
    def get_bios_sn():
        """
        获取BIOS序列号
        :return: BIOS序列号
        """
        c = wmi.WMI()
        for bios_id in c.Win32_BIOS():
            # print(bios_id.SerialNumber.strip)
            return bios_id.SerialNumber.strip()

    @staticmethod
    def get_disk_sn():
        """
        获取硬盘序列号
        :return: 硬盘序列号列表
        """
        c = wmi.WMI()

        disk_sn_list = []
        for physical_disk in c.Win32_DiskDrive():
            # print(physical_disk.SerialNumber)
            # print(physical_disk.SerialNumber.replace(" ", ""))
            disk_sn_list.append(physical_disk.SerialNumber.replace(" ", ""))
        return disk_sn_list
    
def get_unique_id():

    # 组合信息生成唯一ID

    cpu_sn = Hardware.get_cpu_sn()
    # baseboard_sn = Hardware.get_baseboard_sn()
    bios_sn = Hardware.get_bios_sn()
    disk_sn = Hardware.get_disk_sn()
    disk_sn1 = disk_sn[0].replace("_", "")
    disk_sn2 = disk_sn1.replace(".", "")

    register_str = f"{cpu_sn}{bios_sn}{disk_sn2}"

    # print(unique_id)
    # print("CPU序列号：{}".format(Hardware.get_cpu_sn()))
    # print("主板序列号：{}".format(Hardware.get_baseboard_sn()))
    # print("Bios序列号：{}".format(Hardware.get_bios_sn()))
    # print("硬盘序列号：{}".format(Hardware.get_disk_sn()))

    # 打乱
    # chars = list(str)
    # random.shuffle(chars)

    # return ''.join(chars)
    return register_str
