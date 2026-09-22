# 绘制箭头
def draw_arrow(ax, x0, y0, dir, x_scale):
    '''
    x0, y0 箭头长度的中心定位点的坐标
    x1, y1 箭头坐标
    x2, y2 箭尾坐标
    dir 箭头朝向
    '''
    x1 = x0-3*x_scale if dir == 'left' else x0+3*x_scale
    x2 = x0+3*x_scale if dir == 'left' else x0-3*x_scale
    y1 = y0
    y2 = y0
    ax.annotate('', xy=(x1, y1), xytext=(x2, y2), arrowprops=dict(arrowstyle='->', color='red', lw=1.5))


# 绘制弹性支座
def draw_elastic_support(ax, start_x, start_y, x_scale, y_scale):
    # 折线坐标
    x_lst1 = [0, 1, 1.5, 2, 2.5, 3, 4]
    y_lst1 = [0, 0, -0.3, 0, 0.3, 0, 0]
    x_lst1 = [i*0.3*x_scale+start_x for i in x_lst1]
    y_lst1 = [i*0.3*y_scale+start_y for i in y_lst1]
    ax.plot(x_lst1, y_lst1, linestyle='-', color='green', linewidth=0.7) # 折线
    # 直线段坐标
    x_lst2 = [4, 4]
    y_lst2 = [-0.5, 0.5]
    x_lst2 = [i*0.3*x_scale+start_x for i in x_lst2]
    y_lst2 = [i*0.3*y_scale+start_y for i in y_lst2]
    ax.plot(x_lst2, y_lst2, linestyle='-', color='green', linewidth=0.7) # 直线段


# 绘制标高图标
def draw_level_mark(ax, start_x, start_y, txt, x_scale, y_scale):
    # 三角形坐标
    x_lst1 = [2, 0, 4, 2]
    y_lst1 = [0, 0.4, 0.4, 0]
    x_lst1 = [i*0.3*x_scale+start_x for i in x_lst1]
    y_lst1 = [i*0.3*y_scale+start_y for i in y_lst1]
    ax.plot(x_lst1, y_lst1, linestyle='-', color='green', linewidth=0.7) # 三角形
    # 直线段坐标
    x_lst2 = [0, 22]
    y_lst2 = [0, 0]
    x_lst2 = [i*0.3*x_scale+start_x for i in x_lst2]
    y_lst2 = [i*0.3*y_scale+start_y for i in y_lst2]
    ax.plot(x_lst2, y_lst2, linestyle='-', color='green', linewidth=0.7) # 直线段
    # 文字
    x = 5*0.3*x_scale+start_x
    y = 0*0.3*y_scale+start_y
    ax.text(x, y, txt, color='green', fontsize=7, ha='left', va='bottom')


def draw_insert_text(ax, x_scale, y_scale, xlst, ylst, text_Ha):
    y_offset = 0.2*y_scale  # 偏移量，根据你的数据尺度调整
    x_offset = 1*x_scale  # 偏移量，根据你的数据尺度调整
    ny = -1
    for i in range(len(ylst)):
        text_y1 = ylst[i]
        text_x1 = xlst[i]
        if text_x1 != 0:
            text_Va = 'center'
            ny = -1*ny if i != len(ylst)-1 else 1
            y_text = text_y1 + ny*y_offset 
            x_text = text_x1 + x_offset if text_Ha == 'left' else -text_x1 - x_offset
            ax.text(x_text, y_text, f'{text_x1} kPa', ha=text_Ha, va=text_Va, fontsize=7)


def draw_elevation_marker(ax, x, y, text,
                          text_side='right',      # 'left' | 'right'
                          marker_side='above',    # 'above' | 'below'
                          scale=1.0,
                          color='#2b6fd4',
                          fontsize=9,
                          fontweight='normal',
                          facecolor='white',
                          edgewidth=1.0,
                          alpha=1.0,
                          text_voffset=0.0,
                          zorder=15):
    """在 matplotlib 坐标轴上绘制标高标注（等腰三角形 + 三角尖水平线段 + 文字）

    Parameters
    ----------
    ax : matplotlib.axes.Axes
    x, y : float — 三角尖指向的参照坐标（标高位置）
    text : str — 标注文字
    text_side : 'left' | 'right' — 文字在三角的哪一侧
    marker_side : 'above' | 'below' — 三角朝向
        'above': 三角底边在上，三角尖朝下，文字在上方
        'below': 三角底边在下，三角尖朝上，文字在下方
    scale : float — 整体放大系数（同比例缩放三角形和间距）
    color : str — 线条/文字颜色
    fontsize : int/float — 字号
    fontweight : str — 字重 ('normal' / 'bold')
    facecolor : str — 三角填充色
    edgewidth : float — 线条宽
    alpha : float — 透明度
    text_voffset : float — 文字额外纵向偏移（防重叠用）
    zorder : int — 图层顺序
    """
    tri_w = 0.25 * scale
    tri_h = 0.4 * scale
    text_gap = 0.1 * scale
    text_offset = 0.5 * scale

    if marker_side == 'above':
        # 三角尖朝下：底边在 y+h，尖在 y
        tri_x = [x - tri_w, x + tri_w, x]
        tri_y = [y + tri_h, y + tri_h, y]
        text_y = y + tri_h + text_offset + text_voffset
    else:
        # 三角尖朝上：底边在 y-h，尖在 y
        tri_x = [x - tri_w, x + tri_w, x]
        tri_y = [y - tri_h, y - tri_h, y]
        text_y = y - tri_h - text_offset + text_voffset

    # 三角尖水平线段（长度 = 三角形底边长，对称于三角尖）
    ax.plot([x - tri_w, x + tri_w], [y, y],
            color=color, lw=edgewidth, zorder=zorder)

    # 填充三角形
    ax.fill(tri_x, tri_y,
            facecolor=facecolor, edgecolor=color,
            lw=edgewidth, alpha=alpha, zorder=zorder)

    # 文字
    if text_side == 'left':
        ax.text(x - tri_w - text_gap, text_y, text,
                fontsize=fontsize, color=color,
                va='center', ha='right',
                weight=fontweight, zorder=zorder)
    else:
        ax.text(x + tri_w + text_gap, text_y, text,
                fontsize=fontsize, color=color,
                va='center', ha='left',
                weight=fontweight, zorder=zorder)