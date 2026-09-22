# 建模 UI 修改记录 — 画图程序待同步清单

> 记录时间：2026-09-02
> 背景：本次对建模 UI（FEM_MidasCivil/Steel_Pipe_Bailey_Trestle_FEA）做了一批修改，画图程序（Drawing_AutoCad/Steel_Pipe_Bailey_Trestle_CAD）当前为旧版，下个版本要换成新版。以下是需要同步到画图程序 UI / Widgets / Tab / Excel_io 的修改点，按主题分类并注明对应建模文件跟函数，避免遗忘。

---

## 1. 禁用波浪力计算（环境荷载）

**建模文件**：`Trestle_Tab_FEM.py`

### 修改点
- `create_wave_page` 中"计算波浪力"复选框 `self.wave_calc_cb` 的 state 改为 `disabled`，且 `self.wave_calc_var` 强制为 False（不支持波浪力计算）。
- 波浪力参数输入区域（周期/波长/波高）随复选框禁用。

### 画图程序待同步
- `Trestle_Tab_CAD.py` 的 `create_wave_page`（若存在）同样禁用波浪力复选框。

---

## 2. 下部结构二级窗口：联结系验证逻辑

**建模文件**：`Trestle_Widgets_FEM.py` → `on_save`

### 修改点
- 不设置联结系（`has_brace=False`）时，跳过所有联结系字段（`brace_*`）的验证：
  - `on_save` 中 SplitEntry 验证循环加 `if not has_brace and key.startswith("brace"): continue`
  - 截面完整性验证循环加 `if not has_brace and key.startswith("brace"): continue`

### 画图程序待同步
- `Trestle_Widgets_CAD.py` 的 `on_save` 逻辑。

---

## 3. 联结系读写逻辑：brace_form="/" 代表不设置联结系

**建模文件**：`Trestle_Widgets_FEM.py`

### 修改点
- `has_brace` 初始化根据 `brace_form` 推断：
  ```python
  self.vars["has_brace"] = tk.BooleanVar(value=self.data.get("brace_form") not in ("", "/"))
  ```
- `_toggle_brace`：
  - 取消勾选（`has_brace=False`）：`brace_form` 设为 `"/"`
  - 勾选（`has_brace=True`）：若 `brace_form` 为 `/` 则恢复为 `"-"`
- `on_save`：`has_brace=False` 时所有联结系字段（`brace_form/brace_sec/brace_sec_tilt/brace_space/brace_height/brace_material`）保存为 `"/"`。

### 画图程序待同步
- `Trestle_Widgets_CAD.py` 的 `has_brace` 初始化、`_toggle_brace`、`on_save` 清空逻辑。

---

## 4. 材质信息加入下部结构参数表

**建模文件**：`Trestle_Tab_FEM.py` → `_rebuild_table`（columns_keys / headers）

### 修改点
- `columns_keys` 增加材质字段：`pile_material`、`dist_trans_material`、`dist_long_material`、`brace_material`
- `headers` 增加对应材质列标题
- `refresh_cache` 的 `extra` 列表移除材质字段（现由 columns_keys 自动处理）

### 画图程序待同步
- `Trestle_Tab_CAD.py` 的 `_rebuild_table`（columns_keys / headers）与 `refresh_cache`。

---

## 5. SplitEntry 去掉 placeholder 提示字

**建模文件**：`General/UIHandle.py` → `SplitEntry` 类

### 修改点
- 移除 `_apply_placeholder`、`_on_focus_in`、`_on_focus_out` 方法
- 初始化直接插入默认值，不设置灰色前景色
- `get()` 简化：直接读取两个 Entry 值，不再判断灰色前景色

### 关联修改
- `Trestle_Widgets_FEM.py` → `_toggle_brace` 移除依赖 placeholder 前景色的颜色强制设置逻辑（只保留 enable/disable 状态控制）。

### 画图程序待同步
- `SplitEntry` 类在 `General/UIHandle.py`（全局通用，画图程序复用，无需单独同步）。
- `Trestle_Widgets_CAD.py` 若直接操作 `e1/e2` 前景色，需移除。

---

## 6. 联结系间距第二项统一数据结构

**建模文件**：`Trestle_Widgets_FEM.py` → `on_save`

### 修改点
- 联结系间距 `brace_space` 格式统一为 `"{第一项},{第二项}"`；第二项为空统一为 `"/"`：
  ```python
  if self.data.get("has_brace", True):
      bs = str(self.data.get("brace_space", "")).strip()
      if "," in bs:
          v1, v2 = bs.split(",", 1)
          self.data["brace_space"] = f"{v1.strip()},{v2.strip()}" if v2.strip() else f"{v1.strip()},/"
      elif bs:
          self.data["brace_space"] = f"{bs},/"
      else:
          self.data["brace_space"] = ",/"
  ```
- 用例：`"620,"`→`"620,/"`、`"620"`→`"620,/"`、空→`",/"`、`"620,0"` 保持 `"620,0"`。

### 画图程序待同步
- `Trestle_Widgets_CAD.py` 的 `on_save` 处理 `brace_space` 的逻辑。

---

## 7. 默认保存时正常写入材质（关键 bug 修复）

**建模文件**：`Trestle_Widgets_FEM.py` → `_add_mat_combo`

### 修改点
- 初始化时若 `self.data[key]` 无有效牌号，默认设置第一个牌号的同时**写入 `self.data[key]`**：
  ```python
  elif self._steel_brands:
      combo.set(self._steel_brands[0])
      sid = self._steel_spec_id
      self.data[key] = f"{sid},{self._steel_brands[0]}" if sid and self._steel_brands[0] else self._steel_brands[0]
  ```
- 根因：旧版只在用户点选（`<<ComboboxSelected>>`）时才写 `self.data[key]`，导致无联结系→有联结系切换后保存，材质仍未写入。

### 画图程序待同步
- `Trestle_Widgets_CAD.py` 的 `_add_mat_combo`。

---

## 8. 截面名读取不再从参数推算

**建模文件**：`General/UIHandle.py` → `create_section_selector` 内 `_summary`

### 修改点
- `_summary` 不再根据 `D`/`d` 参数推算截面名，直接使用 `custom_label`/`section_name`：
  ```python
  def _summary(data):
      if not data: return "请选择截面"
      sn = data.get("custom_label","") or data.get("section_name","")
      if sn == "自定义" or not sn:
          return "自定义截面" if sn == "自定义" else "请选择截面"
      return sn if sn != "/" else "请选择截面"
  ```
- 根因：旧版 `_summary` 用 `f"{D}X{d}"` 推算，浮点参数产生 `"377.0X6.0"`；且 `_open` 会把该结果写回 `section_name` 污染缓存。
- `_collect` 中用户新建自定义圆钢管自动生成 `custom_lbl = f"{D}×{d}"` 的逻辑**保留**（新建时合理；载入时已有截面名，不走此分支）。

### 画图程序待同步
- `create_section_selector` 与 `_summary` 在全局 `General/UIHandle.py`，画图程序复用，无需单独同步。
- `Trestle_UI_CAD.py` / `Trestle_Widgets_CAD.py` 若自行实现截面选择逻辑需同步。

---

## 通用说明

- **全局共享文件**：`General/UIHandle.py`（SplitEntry、create_section_selector/_summary）为建模、画图程序共用，改一处两处生效，无需重复同步。
- **画图程序待同步文件**（下个版本换新版时）：
  - `Trestle_UI_CAD.py`
  - `Trestle_Widgets_CAD.py`
  - `Trestle_Tab_CAD.py`
  - `Trestle_Excel_io_CAD.py`
- 同步时，将上述建模文件对应函数/逻辑整体比对移植，注意函数名若不同（如 `TrestleUI_CAD` vs `TrestleUI`）需适配。
