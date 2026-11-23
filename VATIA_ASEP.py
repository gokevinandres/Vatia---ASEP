import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import tkinter as tk
from tkinter import messagebox, ttk
import math
import numpy as np
import tempfile
import json
import cmath
from datetime import datetime
from tkinter import filedialog

try:
    from openpyxl.drawing.image import Image as XLImage
    _HAS_XL_IMG = True
except ImportError:
    _HAS_XL_IMG = False

# =====================================================================
# UTILIDAD GLOBAL — Interpolación de colores
# =====================================================================
def interp_color(c1, c2, t):
    """Interpola linealmente entre dos colores hex (#RRGGBB)."""
    t = max(0.0, min(1.0, t))
    r1, g1, b1 = int(c1[1:3], 16), int(c1[3:5], 16), int(c1[5:7], 16)
    r2, g2, b2 = int(c2[1:3], 16), int(c2[3:5], 16), int(c2[5:7], 16)
    return (f'#{int(r1 + (r2 - r1) * t):02x}'
            f'{int(g1 + (g2 - g1) * t):02x}'
            f'{int(b1 + (b2 - b1) * t):02x}')


# =====================================================================
# LOGO VATIA — 4 estilos
# =====================================================================
def draw_vatia_logo(canvas, cx, cy, size=50, style='nexus',
                     bg_color='#1E3A5F', fg_color='#FFFFFF',
                     accent_color='#0E7490', slack_color='#047857',
                     pv_color='#6D28D9', pq_color='#C2410C',
                     phase=0.0):
    r = size / 2
    lw_main = max(2, int(size / 14))
    lw_thin = max(1, int(size / 26))

    if style == 'nexus':
        hex_pts = []
        for i in range(6):
            ang = 2 * math.pi * i / 6 - math.pi / 2
            hex_pts.extend([cx + r * math.cos(ang), cy + r * math.sin(ang)])
        canvas.create_polygon(hex_pts, fill=bg_color, outline='')
        tipos = [slack_color, pv_color, pq_color, pq_color, pv_color, slack_color]
        nodo_r = size * 0.06
        for i in range(6):
            ang = 2 * math.pi * i / 6 - math.pi / 2
            nx = cx + r * 0.88 * math.cos(ang)
            ny = cy + r * 0.88 * math.sin(ang)
            canvas.create_line(cx, cy, nx, ny, fill=accent_color, width=lw_thin)
            pulse = 0.5 + 0.5 * abs(math.sin((phase + i * 0.15) * 2 * math.pi))
            halo_r = nodo_r * (1.4 + 0.6 * pulse)
            canvas.create_oval(nx - halo_r, ny - halo_r, nx + halo_r, ny + halo_r,
                               fill=accent_color, outline='')
            canvas.create_oval(nx - nodo_r, ny - nodo_r, nx + nodo_r, ny + nodo_r,
                               fill=tipos[i], outline=fg_color,
                               width=max(1, int(size / 40)))
        v_h = r * 1.05; v_w = r * 0.55
        pts = [cx - v_w, cy - v_h / 2, cx, cy + v_h / 2, cx + v_w, cy - v_h / 2]
        canvas.create_line(*pts, fill=fg_color, width=lw_main + 1,
                           capstyle=tk.ROUND, joinstyle=tk.ROUND)

    elif style == 'pulse':
        canvas.create_oval(cx - r, cy - r, cx + r, cy + r,
                           fill=bg_color, outline=accent_color, width=lw_main)
        for k in range(3):
            t = (phase + k / 3.0) % 1.0
            ring_r = t * r * 0.85
            opacity = 1.0 - t
            color = interp_color(bg_color, accent_color, opacity)
            canvas.create_oval(cx - ring_r, cy - ring_r, cx + ring_r, cy + ring_r,
                               outline=color, width=2)
        canvas.create_oval(cx - r * 0.35, cy - r * 0.35,
                           cx + r * 0.35, cy + r * 0.35,
                           fill=bg_color, outline='')
        v_h = r * 0.75; v_w = r * 0.42
        pts = [cx - v_w, cy - v_h / 2, cx, cy + v_h / 2, cx + v_w, cy - v_h / 2]
        canvas.create_line(*pts, fill=fg_color, width=lw_main,
                           capstyle=tk.ROUND, joinstyle=tk.ROUND)

    elif style == 'wave':
        canvas.create_oval(cx - r, cy - r, cx + r, cy + r,
                           fill=bg_color, outline=accent_color, width=lw_thin * 2)
        base_y = cy + r * 0.35
        canvas.create_line(cx - r * 0.8, base_y, cx + r * 0.8, base_y,
                           fill=accent_color, width=lw_thin, dash=(3, 3))
        pts = []; n = 60
        for i in range(n + 1):
            t = i / n
            x = cx - r * 0.82 + r * 1.64 * t
            amp = r * 0.55 * (1 - abs(t - 0.5) * 2) + r * 0.15
            phase_shift = phase * 2 * math.pi
            y = cy + r * 0.05 * math.sin(t * 4 * math.pi + phase_shift) - amp * 0.2
            v_shape = (abs(t - 0.5) * 2) * r * 0.35
            y = base_y - v_shape - amp * 0.3 * abs(math.sin(t * math.pi))
            pts.extend([x, y])
        canvas.create_line(*pts, fill=fg_color, width=lw_main,
                           capstyle=tk.ROUND, smooth=True)
        t_p = phase
        x_p = cx - r * 0.82 + r * 1.64 * t_p
        v_shape = (abs(t_p - 0.5) * 2) * r * 0.35
        amp = r * 0.55 * (1 - abs(t_p - 0.5) * 2) + r * 0.15
        y_p = base_y - v_shape - amp * 0.3 * abs(math.sin(t_p * math.pi))
        canvas.create_oval(x_p - size * 0.06, y_p - size * 0.06,
                           x_p + size * 0.06, y_p + size * 0.06,
                           fill=accent_color, outline=fg_color,
                           width=max(1, int(size / 30)))

    elif style == 'tower':
        canvas.create_oval(cx - r, cy - r, cx + r, cy + r,
                           fill=bg_color, outline=accent_color, width=lw_main)
        top_y = cy - r * 0.55; bot_y = cy + r * 0.65
        top_w = r * 0.55; bot_w = r * 0.18
        canvas.create_line(cx - top_w, top_y, cx - bot_w, bot_y,
                           fill=fg_color, width=lw_main, capstyle=tk.ROUND)
        canvas.create_line(cx + top_w, top_y, cx + bot_w, bot_y,
                           fill=fg_color, width=lw_main, capstyle=tk.ROUND)
        for t in [0.15, 0.4, 0.65, 0.88]:
            y_t = top_y + (bot_y - top_y) * t
            w_t = top_w + (bot_w - top_w) * t
            canvas.create_line(cx - w_t, y_t, cx + w_t, y_t,
                               fill=fg_color, width=lw_thin)
        for i in range(3):
            t1 = 0.15 + i * 0.25; t2 = t1 + 0.25
            y1 = top_y + (bot_y - top_y) * t1
            y2 = top_y + (bot_y - top_y) * t2
            w1 = top_w + (bot_w - top_w) * t1
            w2 = top_w + (bot_w - top_w) * t2
            canvas.create_line(cx - w1, y1, cx + w2, y2,
                               fill=accent_color, width=lw_thin)
            canvas.create_line(cx + w1, y1, cx - w2, y2,
                               fill=accent_color, width=lw_thin)
        canvas.create_line(cx - top_w * 1.3, top_y - r * 0.05,
                           cx + top_w * 1.3, top_y - r * 0.05,
                           fill=fg_color, width=lw_main, capstyle=tk.ROUND)
        for sx in (-1, 1):
            for hy in (-r * 0.05, r * 0.08):
                canvas.create_oval(cx + sx * top_w * 1.3 - size * 0.04,
                                   top_y + hy - size * 0.04,
                                   cx + sx * top_w * 1.3 + size * 0.04,
                                   top_y + hy + size * 0.04,
                                   fill=accent_color, outline=fg_color, width=1)
        pulse = 0.5 + 0.5 * abs(math.sin(phase * 2 * math.pi))
        spark_r = size * 0.06 * (1 + 0.5 * pulse)
        canvas.create_oval(cx - spark_r, top_y - r * 0.18 - spark_r,
                           cx + spark_r, top_y - r * 0.18 + spark_r,
                           fill=accent_color, outline=fg_color,
                           width=max(1, int(size / 30)))


# =====================================================================
# DIAGRAMA PNG PARA EXCEL
# =====================================================================
def _generar_diagrama_png(n, nodos, lineas, ruta_png):
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
    except ImportError:
        return False, "matplotlib no está instalado"
    try:
        fig, ax = plt.subplots(figsize=(13, 10), dpi=100)
        ax.set_xlim(-1.25, 1.25); ax.set_ylim(-1.15, 1.22)
        ax.set_aspect('equal'); ax.axis('off'); ax.set_facecolor('white')
        for v in np.arange(-1.2, 1.21, 0.1):
            ax.axvline(v, color='#F4F6F7', lw=0.5, zorder=0)
            ax.axhline(v, color='#F4F6F7', lw=0.5, zorder=0)
        R = 0.72
        pos = [(R*np.cos(2*np.pi*i/n - np.pi/2), R*np.sin(2*np.pi*i/n - np.pi/2)) for i in range(n)]
        for line in lineas:
            i, j = line['a'], line['b']
            x1, y1 = pos[i]; x2, y2 = pos[j]
            ax.plot([x1, x2], [y1, y2], color='#2C3E50', lw=2.5, zorder=2)
            G, B = line['G'], line['B']
            mx, my = (x1+x2)/2, (y1+y2)/2
            lbl = f"Y={G:+.3f}{'+' if B >= 0 else '-'}j{abs(B):.3f}"
            ax.text(mx, my, lbl, fontsize=7, ha='center', va='center',
                    bbox=dict(boxstyle='round,pad=0.25', facecolor='white',
                              edgecolor='#BF360C', lw=0.8), zorder=5, family='monospace')
        for i, (x, y) in enumerate(pos):
            tipo = nodos[i]['tipo']
            color = {3:'#2E7D32', 2:'#6A1B9A', 1:'#E65100'}[tipo]
            ax.plot([x-0.055, x+0.055], [y, y], color='#2D3436', lw=7, solid_capstyle='round', zorder=3)
            ax.scatter([x], [y], s=330, c=color, edgecolors='white', linewidths=2, zorder=4)
            ax.text(x, y, str(i+1), color='white', fontsize=10, fontweight='bold', ha='center', va='center', zorder=5)
            if tipo in (2, 3):
                gy = y + 0.11
                ax.plot([x, x], [y+0.02, gy-0.03], color='#2D3436', lw=1.5, zorder=2)
                ax.add_patch(mpatches.Circle((x, gy), 0.035, facecolor='white', edgecolor='#2E7D32', lw=2, zorder=4))
            if tipo == 1:
                ly = y - 0.10
                ax.plot([x, x], [y-0.02, ly+0.015], color='#2D3436', lw=1.5, zorder=2)
                ax.annotate('', xy=(x, ly-0.025), xytext=(x, ly+0.02),
                            arrowprops=dict(arrowstyle='-|>', color='#E65100', lw=2, mutation_scale=18), zorder=4)
            d = np.hypot(x, y) or 1
            ux, uy = x/d, y/d
            tx, ty = x + ux*0.22, y + uy*0.20
            tipo_txt = {3:'Slack', 2:'PV', 1:'PQ'}[tipo]
            info = f"Nodo {i+1} [{tipo_txt}]\nV={nodos[i]['V']:.3f} p.u.\nθ={nodos[i]['theta']:.2f}°"
            ax.text(tx, ty, info, fontsize=6.5, ha='center', va='center',
                    bbox=dict(boxstyle='round,pad=0.35', facecolor='#FDFEFE', edgecolor=color, lw=1.2),
                    zorder=6, family='monospace')
        ax.text(0, 1.19, 'DIAGRAMA UNIFILAR DEL SISTEMA', fontsize=15, fontweight='bold',
                ha='center', color='#2D3436')
        ax.text(0, 1.12, f'{n} nodos  •  {len(lineas)} línea(s)', fontsize=9,
                ha='center', color='#7B7B7B', style='italic')
        handles = [mpatches.Patch(facecolor='#2E7D32', edgecolor='white', label='Slack'),
                   mpatches.Patch(facecolor='#6A1B9A', edgecolor='white', label='PV'),
                   mpatches.Patch(facecolor='#E65100', edgecolor='white', label='PQ')]
        ax.legend(handles=handles, loc='lower left', fontsize=8, frameon=True, fancybox=True)
        plt.tight_layout()
        plt.savefig(ruta_png, dpi=100, bbox_inches='tight', facecolor='white')
        plt.close(fig)
        return True, "OK"
    except Exception as e:
        return False, str(e)


# =====================================================================
# HOJA DE ITERACIÓN (Excel)
# =====================================================================
def _crear_hoja_iteracion(ws, num_iter, datos, num_nodos,
                          font_titulo, font_sec, font_bold, font_normal,
                          font_mono, fill_titulo, fill_sec, fill_head,
                          fill_data, fill_calc, fill_form, fill_err,
                          border_box, align_c, align_l):
    n = datos['n']; nodos = datos['nodos']; idx_pq = datos['idx_pq']
    Ybus = datos['Ybus']; V_actual = datos['V']; theta_actual = datos['theta']
    P_calc = datos['P_calc']; Q_calc = datos['Q_calc']
    deltaP = datos['deltaP']; deltaQ = datos['deltaQ']
    def sc(r, c, val, font=None, fill=None, border=True, align=None):
        cell = ws.cell(row=r, column=c, value=val)
        if font: cell.font = font
        if fill: cell.fill = fill
        if border: cell.border = border_box
        if align: cell.alignment = align
        return cell
    ws.merge_cells("A1:L1")
    sc(1, 1, f"ITERACIÓN {num_iter} - RESULTADOS DETALLADOS", font_titulo, fill_titulo, align_c)
    ws.row_dimensions[1].height = 24
    r = 3
    ws.merge_cells(f"A{r}:L{r}")
    sc(r, 1, f"A. ESTADO DE LOS NODOS EN LA ITERACIÓN {num_iter}", font_sec, fill_sec, align_c); r += 1
    if num_iter == 0:
        headers = ["Nodo", "Tipo", "V (p.u.)", "θ (rad)", "P_calc", "Q_calc", "ΔP", "ΔQ"]
    else:
        headers = ["Nodo", "Tipo", "V prev", "θ prev (rad)", "V nuevo", "θ nuevo (rad)",
                   "Δθ usado (rad)", "ΔV usado", "P_calc", "Q_calc", "ΔP", "ΔQ"]
    for c, h in enumerate(headers, 1):
        sc(r, c, h, font_bold, fill_head, align_c)
    r += 1
    for i in range(n):
        tipo_str = {1:"PQ", 2:"PV", 3:"Slack"}.get(nodos[i]['tipo'], "PQ")
        dp = deltaP[i] if deltaP[i] is not None else None
        dq = deltaQ[i] if deltaQ[i] is not None else None
        if num_iter == 0:
            fila = [i+1, tipo_str, round(V_actual[i],6), round(theta_actual[i],6),
                    round(P_calc[i],6), round(Q_calc[i],6),
                    round(dp,10) if dp is not None else "---",
                    round(dq,10) if dq is not None else "---"]
        else:
            fila = [i+1, tipo_str, round(datos['V_prev'][i],6), round(datos['theta_prev'][i],6),
                    round(V_actual[i],6), round(theta_actual[i],6),
                    round(datos['delta_theta_usado'][i],6), round(datos['delta_V_usado'][i],6),
                    round(P_calc[i],6), round(Q_calc[i],6),
                    round(dp,10) if dp is not None else "---",
                    round(dq,10) if dq is not None else "---"]
        for c, v in enumerate(fila, 1):
            f = None
            if num_iter == 0:
                if c == 7 and isinstance(dp,(int,float)) and abs(dp) > 0.01: f = fill_err
                elif c == 8 and isinstance(dq,(int,float)) and abs(dq) > 0.01: f = fill_err
            else:
                if c == 11 and isinstance(dp,(int,float)) and abs(dp) > 0.01: f = fill_err
                elif c == 12 and isinstance(dq,(int,float)) and abs(dq) > 0.01: f = fill_err
            sc(r, c, v, font_normal, f, align_c)
        r += 1
    r += 2
    ws.merge_cells(f"A{r}:L{r}")
    sc(r, 1, "B. MATRIZ JACOBIANA (J) - CON FÓRMULAS", font_sec, fill_sec, align_c); r += 1
    for f in [
        "J = [ H  N ]   donde H = ∂P/∂θ,  N = ∂P/∂V,  M = ∂Q/∂θ,  L = ∂Q/∂V",
        "    [ M  L ]",
        "H_ii = -Q_i - B_ii·V_i²       H_ij = V_i·V_j·(G_ij·sin(θi-θj) - B_ij·cos(θi-θj))",
        "N_ii = P_i/V_i + G_ii·V_i      N_ij = V_i·(G_ij·cos(θi-θj) + B_ij·sin(θi-θj))",
        "M_ii = P_i - G_ii·V_i²         M_ij = -V_i·V_j·(G_ij·cos(θi-θj) + B_ij·sin(θi-θj))",
        "L_ii = Q_i/V_i - B_ii·V_i      L_ij = V_i·(G_ij·sin(θi-θj) - B_ij·cos(θi-θj))",
    ]:
        ws.merge_cells(f"A{r}:L{r}"); sc(r, 1, f, font_mono, fill_form, align_l); r += 1
    J = datos['J']; m = J.shape[0]
    var_names = []
    for i in range(n):
        if nodos[i]['tipo'] != 3: var_names.append(f"θ{i+1}")
    for i in idx_pq: var_names.append(f"V{i+1}")
    row_names = []
    for i in range(n):
        if nodos[i]['tipo'] != 3: row_names.append(f"ΔP{i+1}")
    for i in idx_pq: row_names.append(f"ΔQ{i+1}")
    sc(r, 1, "Ec./Var", font_bold, fill_head, align_c)
    for c, name in enumerate(var_names, 2):
        sc(r, c, name, font_bold, fill_head, align_c)
    sc(r, 2+len(var_names), "Mismatch", font_bold, fill_head, align_c)
    sc(r, 3+len(var_names), "Δx (solución)", font_bold, fill_head, align_c)
    r += 1
    if num_iter == 0:
        mismatch = datos['mismatch']
    else:
        mismatch = datos.get('mismatch_actual', datos.get('mismatch_usado'))
        if mismatch is None:
            tmp = []
            for i in range(n):
                if nodos[i]['tipo'] != 3: tmp.append(deltaP[i] if deltaP[i] is not None else 0.0)
            for i in idx_pq: tmp.append(deltaQ[i] if deltaQ[i] is not None else 0.0)
            mismatch = np.array(tmp)
    delta_x = datos.get('delta_x', np.zeros(m))
    for ri in range(m):
        sc(r, 1, row_names[ri], font_bold, fill_head, align_c)
        for c in range(m):
            val = J[ri, c]
            txt = f"{val:.6f}" if abs(val) > 1e-6 else "0.000000"
            sc(r, c+2, txt, font_mono, fill_data, align_c)
        mis = mismatch[ri] if ri < len(mismatch) else 0
        sc(r, 2+len(var_names), f"{mis:.10f}", font_mono, fill_calc, align_c)
        dx = delta_x[ri] if ri < len(delta_x) else 0
        sc(r, 3+len(var_names), f"{dx:.10f}", font_mono, fill_calc, align_c)
        r += 1
    r += 2
    ws.merge_cells(f"A{r}:L{r}")
    sc(r, 1, "C. SISTEMA LINEAL: J · Δx = mismatch", font_sec, fill_sec, align_c); r += 1
    for f in [
        "Se resuelve el sistema lineal:   J(θ,V) · Δx = mismatch(θ,V)",
        "donde Δx = [Δθ₁... Δθₙ  ΔV₁... ΔVₚ]ᵀ  (variables de control)",
        "mismatch = [ΔP₁... ΔPₙ  ΔQ₁... ΔQₚ]ᵀ  (P_esp - P_calc;  Q_esp - Q_calc)",
    ]:
        ws.merge_cells(f"A{r}:L{r}"); sc(r, 1, f, font_mono, fill_form, align_l); r += 1
    r += 1
    sc(r, 1, "Variable", font_bold, fill_head, align_c)
    sc(r, 2, "Δx (solución)", font_bold, fill_head, align_c); r += 1
    var_names_delta = []
    for i in range(n):
        if nodos[i]['tipo'] != 3: var_names_delta.append(f"Δθ{i+1} (rad)")
    for i in idx_pq: var_names_delta.append(f"ΔV{i+1} (p.u.)")
    for k, name in enumerate(var_names_delta):
        sc(r, 1, name, font_bold, fill_data, align_c)
        val = delta_x[k] if k < len(delta_x) else 0.0
        sc(r, 2, f"{val:.10f}", font_mono, fill_calc, align_c); r += 1
    r += 2
    ws.merge_cells(f"A{r}:L{r}")
    sc(r, 1, "D. ACTUALIZACIÓN DE VARIABLES", font_sec, fill_sec, align_c); r += 1
    for c, h in enumerate(["Variable","Valor anterior","Corrección","Valor nuevo","Verificación"], 1):
        sc(r, c, h, font_bold, fill_head, align_c)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7); r += 1
    if num_iter == 0:
        V_prev = datos['V']; theta_prev = datos['theta']
        d_theta = datos['delta_theta']; d_V = datos['delta_V']
    else:
        V_prev = datos['V_prev']; theta_prev = datos['theta_prev']
        d_theta = datos['delta_theta_usado']; d_V = datos['delta_V_usado']
    theta_new = theta_prev + d_theta; V_new = V_prev + d_V
    for i in range(n):
        if nodos[i]['tipo'] != 3:
            sc(r, 1, f"θ{i+1}", font_bold, fill_data, align_c)
            sc(r, 2, f"{theta_prev[i]:.6f}", font_mono, fill_data, align_c)
            sc(r, 3, f"{d_theta[i]:.10f}", font_mono, fill_calc, align_c)
            sc(r, 4, f"{theta_new[i]:.6f}", font_mono, fill_data, align_c)
            ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
            sc(r, 5, f"{theta_prev[i]:.4f} + {d_theta[i]:.6f} = {theta_new[i]:.4f}", font_mono, fill_form, align_l); r += 1
        if nodos[i]['tipo'] == 1:
            sc(r, 1, f"V{i+1}", font_bold, fill_data, align_c)
            sc(r, 2, f"{V_prev[i]:.6f}", font_mono, fill_data, align_c)
            sc(r, 3, f"{d_V[i]:.10f}", font_mono, fill_calc, align_c)
            sc(r, 4, f"{V_new[i]:.6f}", font_mono, fill_data, align_c)
            ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
            sc(r, 5, f"{V_prev[i]:.4f} + {d_V[i]:.6f} = {V_new[i]:.4f}", font_mono, fill_form, align_l); r += 1
    r += 2
    ws.merge_cells(f"A{r}:L{r}")
    sc(r, 1, "E. MATRIZ Ybus UTILIZADA", font_sec, fill_sec, align_c); r += 1
    sc(r, 1, "", font_bold, fill_head, align_c)
    for j in range(n): sc(r, j+2, f"N{j+1}", font_bold, fill_head, align_c)
    r += 1
    for i in range(n):
        sc(r, 1, f"N{i+1}", font_bold, fill_head, align_c)
        for j in range(n):
            val = Ybus[i][j]
            if abs(val.imag) < 1e-12: txt = f"{val.real:.4f}"
            elif abs(val.real) < 1e-12: txt = f"{val.imag:.4f}j"
            else:
                s = "+" if val.imag >= 0 else "-"
                txt = f"{val.real:.4f}{s}{abs(val.imag):.4f}j"
            sc(r, j+2, txt, font_mono, fill_data, align_c)
        r += 1
    ws.column_dimensions['A'].width = 16
    for c in "BCDEFGHIJKL": ws.column_dimensions[c].width = 16


# =====================================================================
# GENERADOR EXCEL
# =====================================================================
def crear_excel_completo(num_nodos, datos_nodos_user, datos_lineas_user,
                          datos_iter, nodos_info, lineas_info):
    try:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        ruta_archivo = os.path.join(desktop, f"Flujo_Potencia_{num_nodos}_Nodos.xlsx")
        wb = openpyxl.Workbook()
        font_titulo = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        font_sec = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        font_bold = Font(name="Calibri", size=10, bold=True)
        font_normal = Font(name="Calibri", size=10)
        font_mono = Font(name="Consolas", size=9)
        fill_titulo = PatternFill("solid", start_color="2D3436")
        fill_sec = PatternFill("solid", start_color="BF360C")
        fill_head = PatternFill("solid", start_color="EFEBE0")
        fill_data = PatternFill("solid", start_color="FAF8F3")
        fill_calc = PatternFill("solid", start_color="FFF3E0")
        fill_form = PatternFill("solid", start_color="FDF3E7")
        fill_err = PatternFill("solid", start_color="F8D7DA")
        border_thin = Side(style="thin", color="000000")
        border_box = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
        align_c = Alignment(horizontal="center", vertical="center")
        align_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if 0 in datos_iter: Ybus = datos_iter[0]['Ybus']
        else:
            Ybus = [[0j]*num_nodos for _ in range(num_nodos)]
            for line in lineas_info:
                a, b = line['a'], line['b']
                G, B = line['G'], line['B']; ysh = line.get('y_shunt', 0.0)
                y_serie = complex(G, B)
                Ybus[a][a] += y_serie + complex(0, ysh); Ybus[b][b] += y_serie + complex(0, ysh)
                Ybus[a][b] -= y_serie; Ybus[b][a] -= y_serie
        def fmt_c(z):
            if abs(z.imag) < 1e-12: return f"{z.real:.6f}"
            if abs(z.real) < 1e-12: return f"{z.imag:.6f}j"
            s = "+" if z.imag >= 0 else "-"
            return f"{z.real:.6f} {s} {abs(z.imag):.6f}j"
        ws1 = wb.active; ws1.title = "Datos y Diagrama"
        def sc1(r, c, val, font=None, fill=None, border=True, align=None):
            cell = ws1.cell(row=r, column=c, value=val)
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border_box
            if align: cell.alignment = align
            return cell
        ws1.merge_cells("A1:H1")
        sc1(1, 1, f"SISTEMA DE POTENCIA DE {num_nodos} NODOS - FLUJO NEWTON-RAPHSON",
            font_titulo, fill_titulo, align=align_c)
        ws1.row_dimensions[1].height = 24
        r = 3
        ws1.merge_cells(f"A{r}:H{r}")
        sc1(r, 1, "1. DATOS NODALES", font_sec, fill_sec, align=align_c); r += 1
        for c, h in enumerate(["Nodo","PGi (MW)","QGi (MVAr)","PDi (MW)","QDi (MVAr)","Ei (kV)","Ebase (kV)","Tipo"], 1):
            sc1(r, c, h, font_bold, fill_head, align=align_c)
        r += 1
        for row_data in datos_nodos_user:
            for c, val in enumerate(row_data[:8], 1):
                sc1(r, c, val, font_normal, fill_data, align=align_c)
            r += 1
        r += 2
        ws1.merge_cells(f"A{r}:E{r}")
        sc1(r, 1, "2. DATOS DE LÍNEAS", font_sec, fill_sec, align=align_c); r += 1
        for c, h in enumerate(["Nodo A","Nodo B","G (p.u.)","B (p.u.)","y_shunt/2 (p.u.)"], 1):
            sc1(r, c, h, font_bold, fill_head, align=align_c)
        r += 1
        for row_data in datos_lineas_user:
            for c, val in enumerate(row_data[:5], 1):
                sc1(r, c, val, font_normal, fill_data, align=align_c)
            r += 1
        r += 2
        ws1.merge_cells(f"A{r}:H{r}")
        sc1(r, 1, "3. ITERACIÓN 0", font_sec, fill_sec, align=align_c); r += 1
        if 0 in datos_iter:
            d0 = datos_iter[0]
            for c, h in enumerate(["Nodo","Tipo","V (p.u.)","θ (rad)","P_calc","Q_calc","ΔP","ΔQ"], 1):
                sc1(r, c, h, font_bold, fill_head, align=align_c)
            r += 1
            for i in range(num_nodos):
                tipo_str = {1:"PQ", 2:"PV", 3:"Slack"}[d0['nodos'][i]['tipo']]
                dp = d0['deltaP'][i] if d0['deltaP'][i] is not None else "---"
                dq = d0['deltaQ'][i] if d0['deltaQ'][i] is not None else "---"
                fila = [i+1, tipo_str, round(d0['V'][i],6), round(d0['theta'][i],6),
                        round(d0['P_calc'][i],6), round(d0['Q_calc'][i],6),
                        round(dp,10) if isinstance(dp,(int,float)) else dp,
                        round(dq,10) if isinstance(dq,(int,float)) else dq]
                for c, v in enumerate(fila, 1):
                    f = None
                    if c == 7 and isinstance(dp,(int,float)) and abs(dp) > 0.01: f = fill_err
                    elif c == 8 and isinstance(dq,(int,float)) and abs(dq) > 0.01: f = fill_err
                    sc1(r, c, v, font_normal, f, align=align_c)
                r += 1
        r += 2
        ws1.merge_cells(f"A{r}:H{r}")
        sc1(r, 1, "4. DIAGRAMA UNIFILAR", font_sec, fill_sec, align=align_c)
        r += 2
        try:
            ruta_png = os.path.join(tempfile.gettempdir(), f"diagrama_flujo_{num_nodos}.png")
            ok, msg = _generar_diagrama_png(num_nodos, nodos_info, lineas_info, ruta_png)
            if ok and _HAS_XL_IMG:
                try:
                    img = XLImage(ruta_png); img.width = 720; img.height = 555
                    ws1.add_image(img, f"A{r}")
                except Exception as e_img:
                    sc1(r, 1, f"[Imagen no insertada: {e_img}]", font_normal, fill_calc, border=False)
        except Exception as e_diag:
            sc1(r, 1, f"[Error: {e_diag}]", font_normal, fill_calc, border=False)
        ws1.column_dimensions['A'].width = 14
        for c in "BCDEFGH": ws1.column_dimensions[c].width = 18
        ws2 = wb.create_sheet("Matriz Ybus")
        def sc2(r, c, val, font=None, fill=None, border=True, align=None):
            cell = ws2.cell(row=r, column=c, value=val)
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border_box
            if align: cell.alignment = align
            return cell
        ws2.merge_cells("A1:H1")
        sc2(1, 1, f"MATRIZ Ybus ({num_nodos}x{num_nodos})", font_titulo, fill_titulo, align=align_c)
        ws2.row_dimensions[1].height = 24
        r2 = 3
        sc2(r2, 1, "", font_bold, fill_head, align=align_c)
        for j in range(num_nodos): sc2(r2, j+2, f"Nodo {j+1}", font_bold, fill_head, align=align_c)
        r2 += 1
        for i in range(num_nodos):
            sc2(r2, 1, f"Nodo {i+1}", font_bold, fill_head, align=align_c)
            for j in range(num_nodos):
                sc2(r2, j+2, fmt_c(Ybus[i][j]), font_normal, fill_data, align=align_c)
            r2 += 1
        ws2.column_dimensions['A'].width = 12
        for c in "BCDEFG": ws2.column_dimensions[c].width = 24
        ws2.column_dimensions['H'].width = 22
        if datos_iter:
            for num_iter in sorted(datos_iter.keys()):
                ws_it = wb.create_sheet(f"Iteración {num_iter}")
                _crear_hoja_iteracion(ws_it, num_iter, datos_iter[num_iter], num_nodos,
                    font_titulo, font_sec, font_bold, font_normal, font_mono,
                    fill_titulo, fill_sec, fill_head, fill_data, fill_calc,
                    fill_form, fill_err, border_box, align_c, align_l)
        wb.save(ruta_archivo)
        try: os.startfile(ruta_archivo)
        except Exception: pass
        messagebox.showinfo("Éxito",
            f"Excel generado en:\n{ruta_archivo}\n\n"
            f"• Hoja 1: Datos + Iter 0 + Diagrama\n"
            f"• Hoja 2: Ybus\n"
            f"• Hojas 3+: {len(datos_iter)} iteración(es)")
    except Exception as e:
        import traceback
        messagebox.showerror("Error", f"Error al generar Excel:\n{e}\n\n{traceback.format_exc()}")


# =====================================================================
# TIPOGRAFÍA E ICONOGRAFÍA GLOBAL
# =====================================================================
FONT_FAMILY     = "Segoe UI"
FONT_FAMILY_ALT = "Segoe UI"
FONT_MONO       = "Consolas"
FONT_MONO_ALT   = "Consolas"

def font(size=10, weight="normal", italic=False, family=None):
    fam = family or FONT_FAMILY
    style = "italic" if italic else "normal"
    return (fam, size, style if weight == "normal" else f"{weight} {style}".strip())

def font_m(size=9, weight="normal"):
    return (FONT_MONO, size, weight)

# =====================================================================
# TOOLTIPS
# =====================================================================
class ToolTip:
    def __init__(self, widget, text, delay=500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self._after_id = None
        widget.bind('<Enter>', self._schedule, add='+')
        widget.bind('<Leave>', self._cancel, add='+')
        widget.bind('<ButtonPress>', self._cancel, add='+')

    def _schedule(self, event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay, self._show)

    def _cancel(self, event=None):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None
        self._hide()

    def _show(self):
        if self.tip_window: return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify='left',
                 background='#1F2937', foreground='white',
                 relief='flat', borderwidth=0,
                 font=('Segoe UI', 9), padx=8, pady=4).pack()

    def _hide(self):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None
# =====================================================================
# APLICACIÓN PRINCIPAL
# =====================================================================
class AppFlujoPotencia:
    C = {
        'bg':          '#EDF1F6',
        'panel':       '#FFFFFF',
        'surface':     '#F7F9FC',
        'sidebar':     '#E2E8F0',
        'primary':     '#1E3A5F',
        'primary_l':   '#2C5282',
        'primary_d':   '#152741',
        'accent':      '#0E7490',
        'accent_l':    '#0891B2',
        'slack':       '#047857',
        'pv':          '#6D28D9',
        'pq':          '#C2410C',
        'info':        '#0369A1',
        'amber':       '#B45309',
        'text':        '#1F2937',
        'text_l':      '#4B5563',
        'text_ll':     '#94A3B8',
        'border':      '#CBD5E1',
        'border_d':    '#94A3B8',
        'head_bg':     '#E2E8F0',
        'row_even':    '#FFFFFF',
        'row_odd':     '#F7F9FC',
        'danger':      '#B91C1C',
        'warning':     '#B45309',
        'success':     '#047857',
        'gray_dark':   '#475569',
        'glow':        '#0E7490',
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Vatia — Simulador de Sistemas de Potencia")
        self.root.geometry("1500x900")
        self.root.minsize(1180, 720)
        self.root.configure(bg=self.C['bg'])

        self._set_app_icon()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Listas de widgets
        self.tipo_widgets = []
        self.V_fijo_widgets = []
        self.theta_fijo_widgets = []
        self.Pd_widgets = []
        self.Qd_widgets = []
        self.Pg_widgets = []
        self.Qg_widgets = []
        self.tipo_labels_pot = []
        self.entries_lineas = []
        self.iter0_V_entries = []
        self.iter0_theta_entries = []
        self.iter0_tipo_labels = []

        # Bases del sistema
        self.BASE_MVA = 100.0
        self.V_BASE_KV = 220.0

        # Estado
        self.datos_iteracion = {}
        self.ventanas_resultados = {}
        self._updating = False
        self._last_tipo_values = {}
        self.sidebar_labels = {}

        # IDs de animaciones
        self._logo_after_id = None
        self._sidebar_after_id = None
        self._clock_after_id = None
        self._logo_pulse = 0.0
        self._flow_particles = []
        self._flow_spawn_counter = 0
        self._resultados_pulse = 0.0

        self._build_menu()
        self._setup_styles()
        self._build_header()
        self._build_toolbar()
        self._build_base_bar()
        self._build_main_area()
        self._build_action_bar()
        self._build_status_bar()

        self.root.bind('<F1>', lambda e: self._open_help())

        self.generar_tablas()
        self._set_progress(0.0)

    @staticmethod
    def _parse_float(s, default=0.0):
        try:
            if s is None: return default
            s = str(s).strip().replace(',', '.')
            return float(s) if s else default
        except Exception:
            return default

    @staticmethod
    def _parse_int(s, default=0):
        try:
            if s is None: return default
            s = str(s).strip().replace(',', '.')
            return int(float(s)) if s else default
        except Exception:
            return default

    @staticmethod
    def nodo_label(idx):
        """Devuelve la etiqueta del nodo con subíndice Unicode."""
        sub = str(idx + 1)
        return "_" + sub

    # =================================================================
    # ICONO
    # =================================================================
    def _set_app_icon(self):
        try:
            from PIL import Image, ImageDraw, ImageTk
            ico_path = os.path.join(tempfile.gettempdir(), "vatia_icon.ico")
            if not os.path.exists(ico_path):
                img = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
                draw = ImageDraw.Draw(img)
                pts = []
                for i in range(6):
                    ang = math.pi * i / 3 - math.pi / 2
                    pts.append((128 + 115 * math.cos(ang),
                                128 + 115 * math.sin(ang)))
                draw.polygon(pts, fill=(30, 58, 95, 255))
                pts2 = []
                for i in range(6):
                    ang = math.pi * i / 3 - math.pi / 2
                    pts2.append((128 + 100 * math.cos(ang),
                                 128 + 100 * math.sin(ang)))
                draw.polygon(pts2, outline=(14, 116, 144, 255), width=6)
                draw.line([(72, 90), (128, 172), (184, 90)],
                          fill=(255, 255, 255, 255), width=26, joint='curve')
                draw.ellipse([120, 164, 136, 180], fill=(14, 116, 144, 255))
                img.save(ico_path, sizes=[(16,16), (24,24), (32,32),
                                           (48,48), (64,64), (128,128), (256,256)])
            self.root.iconbitmap(ico_path)
            from PIL import Image as PILImage
            for size in (16, 32, 48, 64, 128, 256):
                im = PILImage.open(ico_path).resize((size, size), PILImage.LANCZOS)
                photo = ImageTk.PhotoImage(im)
                if not hasattr(self, '_icon_refs'):
                    self._icon_refs = []
                self._icon_refs.append(photo)
                try:
                    self.root.wm_iconphoto(False, photo)
                except Exception:
                    pass
        except Exception as e:
            print(f"[Icon] No se pudo aplicar el icono: {e}")

    # =================================================================
    # MENÚ
    # =================================================================
    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        m_arch = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Archivo", menu=m_arch)
        m_arch.add_command(label="Nuevo proyecto\t\tCtrl+N", command=self._nuevo_proyecto)
        m_arch.add_command(label="Abrir proyecto...\tCtrl+O", command=self._abrir_proyecto)
        m_arch.add_command(label="Guardar\t\tCtrl+S", command=self._guardar_proyecto)
        m_arch.add_separator()
        m_arch.add_command(label="Exportar a Excel...\tCtrl+E", command=self.exportar_excel)
        m_arch.add_separator()
        m_arch.add_command(label="Salir", command=self.root.quit)

        m_calc = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Cálculo", menu=m_calc)
        m_calc.add_command(label="Ejecutar iteración 0\tF5", command=self.calcular_iteracion_cero)
        m_calc.add_command(label="Ejecutar iteración 1", command=self.calcular_iteracion_uno)
        m_calc.add_command(label="Ejecutar iteración 2", command=self.calcular_iteracion_dos)
        m_calc.add_command(label="Ejecutar iteración 3", command=self.calcular_iteracion_tres)
        m_calc.add_separator()
        m_calc.add_command(label="Calcular todo (0 → 3)", command=self._calcular_todas)

        m_ver = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ver", menu=m_ver)
        m_ver.add_command(label="Diagrama unifilar\tCtrl+D", command=self.mostrar_diagrama_unifilar)
        m_ver.add_command(label="Resultados finales", command=self.mostrar_resultados_finales)

        m_ayuda = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=m_ayuda)
        m_ayuda.add_command(label="Manual de uso\tF1", command=self._open_help)
        m_ayuda.add_separator()
        m_ayuda.add_command(label="Acerca de Vatia...", command=self._acerca_de)

        self.root.bind('<Control-n>', lambda e: self._nuevo_proyecto())
        self.root.bind('<Control-o>', lambda e: self._abrir_proyecto())
        self.root.bind('<Control-s>', lambda e: self._guardar_proyecto())
        self.root.bind('<Control-e>', lambda e: self.exportar_excel())
        self.root.bind('<Control-d>', lambda e: self.mostrar_diagrama_unifilar())
        self.root.bind('<F5>', lambda e: self.calcular_iteracion_cero())

    # =================================================================
    # ESTILOS
    # =================================================================
    def _setup_styles(self):
        C = self.C
        style = ttk.Style()
        try: style.theme_use('clam')
        except Exception: pass
        style.configure('TNotebook', background=C['bg'], borderwidth=0)
        style.configure('TNotebook.Tab', background=C['head_bg'],
                        foreground=C['text_l'],
                        padding=[20, 10], font=font(10, "bold"))
        style.map('TNotebook.Tab', background=[('selected', C['primary'])],
                  foreground=[('selected', 'white')])
        style.configure('TEntry', padding=5, font=font(9),
                        fieldbackground=C['surface'], foreground=C['text'])
        style.configure('TCombobox', padding=4, font=font(9))
        style.configure('TLabelframe', background=C['panel'], borderwidth=1,
                        relief='solid', bordercolor=C['border'])
        style.configure('TLabelframe.Label', background=C['panel'],
                        foreground=C['primary'], font=font(10, "bold"))

    # =================================================================
    # HEADER
    # =================================================================
    def _build_header(self):
        C = self.C
        hdr = tk.Frame(self.root, bg=C['primary_d'], height=130)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['accent'], height=4).pack(fill=tk.X, side=tk.TOP)

        centro = tk.Frame(hdr, bg=C['primary_d'])
        centro.pack(expand=True, pady=(8, 0))

        self.logo_canvas = tk.Canvas(centro, width=280, height=85,
                                       bg=C['primary_d'], highlightthickness=0)
        self.logo_canvas.pack()

        self._logo_draw_size = 78
        self._logo_cx = 45
        self._logo_cy = 42
        
        left = tk.Frame(hdr, bg=C['primary_d'])
        left.place(relx=0.0, rely=0.5, anchor='w', x=24)
        badge = tk.Frame(left, bg=C['accent'])
        badge.pack()
        tk.Label(badge, text="MÓDULO", font=font(8, "bold"),
                 bg=C['accent'], fg=C['primary_d']).pack(padx=14, pady=(6, 0))
        tk.Label(badge, text="FLUJOS", font=font(13, "bold"),
                 bg=C['accent'], fg='white').pack(padx=14, pady=(0, 6))
        self.root.after(200, self._animate_logo)


    # =================================================================
    # BARRA DE BASES
    # =================================================================
    def _build_base_bar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(0, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['border'])
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel'])
        inner.pack(padx=18, pady=10, fill=tk.X)

        tk.Label(inner, text="DATOS DEL SISTEMA", font=font(9, "bold"),
                 bg=C['panel'], fg=C['primary']).pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(inner, text="Potencia base:", font=font(9),
                 bg=C['panel'], fg=C['text_l']).pack(side=tk.LEFT)
        self.entry_base_mva = ttk.Entry(inner, width=8,
                                          justify='center', font=font(10))
        self.entry_base_mva.pack(side=tk.LEFT, padx=(6, 4), ipady=2)
        self.entry_base_mva.insert(0, "100")
        tk.Label(inner, text="MVA", font=font(9, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=(0, 24))

        tk.Label(inner, text="Voltaje base:", font=font(9),
                 bg=C['panel'], fg=C['text_l']).pack(side=tk.LEFT)
        self.entry_base_kv = ttk.Entry(inner, width=8,
                                         justify='center', font=font(10))
        self.entry_base_kv.pack(side=tk.LEFT, padx=(6, 4), ipady=2)
        self.entry_base_kv.insert(0, "220")
        tk.Label(inner, text="kV", font=font(9, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=(0, 24))

        tk.Label(inner, text="→  Z_base = (V_base)² / S_base",
                 font=font_m(9, "italic"), bg=C['panel'], fg=C['text_ll']
                 ).pack(side=tk.LEFT)

        self.entry_base_mva.bind('<KeyRelease>', lambda e: self._actualizar_bases())
        self.entry_base_kv.bind('<KeyRelease>', lambda e: self._actualizar_bases())

    def _actualizar_bases(self):
        try:
            self.BASE_MVA = self._parse_float(self.entry_base_mva.get(), 100.0)
            if self.BASE_MVA <= 0:
                self.BASE_MVA = 100.0
        except Exception:
            pass
        try:
            self.V_BASE_KV = self._parse_float(self.entry_base_kv.get(), 220.0)
            if self.V_BASE_KV <= 0:
                self.V_BASE_KV = 220.0
        except Exception:
            pass

    # =================================================================
    # TOOLBAR
    # =================================================================
    def _build_toolbar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(10, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['border'])
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel'])
        inner.pack(padx=18, pady=12, fill=tk.X)

        tk.Label(inner, text="Número de Nodos (N):", font=font(10, "bold"),
                 bg=C['panel'], fg=C['text']).pack(side=tk.LEFT)
        self.entry_n = ttk.Entry(inner, width=6, justify='center', font=font(11))
        self.entry_n.pack(side=tk.LEFT, padx=8, ipady=3)
        self.entry_n.insert(0, "3")

        btn = tk.Button(inner, text="  Generar Tablas  ",
                        command=self.generar_tablas, bg=C['primary'], fg='white',
                        font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                        activebackground=C['primary_l'], activeforeground='white',
                        padx=16, pady=7, borderwidth=0)
        btn.pack(side=tk.LEFT, padx=10)

        tk.Label(inner, text="Reabrir:", font=font(9, "bold"),
                 bg=C['panel'], fg=C['text_l']).pack(side=tk.LEFT, padx=(26, 6))

        self.btn_redisplay = {}
        for it in range(4):
            b = tk.Button(inner, text=f"Iter {it}",
                          command=lambda k=it: self.redisplay_iter(k),
                          bg=C['surface'], fg=C['text_l'], font=font(9, "bold"),
                          relief=tk.FLAT, cursor='hand2', padx=12, pady=4,
                          state=tk.DISABLED, borderwidth=0,
                          activebackground=C['amber'], activeforeground=C['primary_d'])
            b.pack(side=tk.LEFT, padx=2)
            self.btn_redisplay[it] = b

        self.btn_redisplay0 = self.btn_redisplay[0]
        self.btn_redisplay1 = self.btn_redisplay[1]
        self.btn_redisplay2 = self.btn_redisplay[2]
        self.btn_redisplay3 = self.btn_redisplay[3]

        self.btn_resultados_finales = tk.Button(
            inner, text="  ★  RESULTADOS FINALES  ",
            command=self.mostrar_resultados_finales,
            bg=C['amber'], fg='white',
            font=font(10, "bold"),
            relief=tk.FLAT, cursor='hand2',
            padx=20, pady=7,
            state=tk.DISABLED, borderwidth=0,
            disabledforeground=C['text_ll'])
        self.btn_resultados_finales.pack(side=tk.RIGHT, padx=(10, 2))

    # =================================================================
    # MAIN AREA
    # =================================================================
    def _build_main_area(self):
        C = self.C
        main = tk.Frame(self.root, bg=C['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        self.paned = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.sidebar = tk.Frame(self.paned, bg=C['sidebar'], width=210)
        self.paned.add(self.sidebar, weight=0)
        self._build_sidebar()

        center = tk.Frame(self.paned, bg=C['bg'])
        self.paned.add(center, weight=5)
        self.notebook = ttk.Notebook(center)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.notebook.bind('<<NotebookTabChanged>>',
                            lambda e: self.root.after(30, self._sincronizar_tipos))

        right = tk.Frame(self.paned, bg=C['bg'], width=380)
        self.paned.add(right, weight=4)
        right_lf = tk.Frame(right, bg=C['panel'], highlightthickness=1,
                             highlightbackground=C['border'])
        right_lf.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        tk.Label(right_lf, text="  VISTA PREVIA UNIFILAR  ",
                 font=font(10, "bold"), bg=C['head_bg'], fg=C['accent'],
                 pady=8).pack(fill=tk.X)

        self.preview_canvas = tk.Canvas(right_lf, bg=C['panel'], highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.preview_canvas.bind('<Configure>', lambda e: self._dibujar_preview())

    # =================================================================
    # SIDEBAR
    # =================================================================
    def _build_sidebar(self):
        C = self.C
        sb = self.sidebar
        sb.pack_propagate(False)
        tk.Label(sb, text="  PANEL DE CONTROL  ",
                 font=font(10, "bold"), bg=C['primary'], fg='white',
                 pady=10).pack(fill=tk.X)
        tk.Frame(sb, bg=C['accent'], height=2).pack(fill=tk.X)

        def _stat(key, label, color=None):
            frame = tk.Frame(sb, bg=C['sidebar'])
            frame.pack(fill=tk.X, padx=14, pady=(12, 4))
            tk.Label(frame, text=label.upper(), font=font(8, "bold"),
                     bg=C['sidebar'], fg=C['text_l']).pack(anchor='w')
            lbl = tk.Label(frame, text="—", font=font(18, "bold"),
                           bg=C['sidebar'], fg=color or C['text'])
            lbl.pack(anchor='w')
            self.sidebar_labels[key] = lbl

        def _sep():
            tk.Frame(sb, bg=C['border'], height=1).pack(fill=tk.X, padx=14, pady=8)

        _stat('nodos', 'Nodos del sistema', C['accent'])
        _stat('lineas', 'Líneas de transmisión', C['info'])
        _sep()
        _stat('slack', 'Nodos Slack', C['slack'])
        _stat('pv', 'Nodos PV', C['pv'])
        _stat('pq', 'Nodos PQ', C['pq'])
        _sep()
        _stat('iter', 'Última iteración', C['text'])
        _stat('estado', 'Estado', C['text_l'])
        _stat('error', 'Error máximo', C['warning'])

        pie = tk.Frame(sb, bg=C['sidebar'])
        pie.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        tk.Label(pie, text="por Kevin", font=font(9, italic=True),
                 bg=C['sidebar'], fg=C['text_l']).pack()
        tk.Label(pie, text="Vatia © 2026", font=font_m(8),
                 bg=C['sidebar'], fg=C['text_ll']).pack()
        self.root.after(400, self._update_sidebar)

    # =================================================================
    # ACTION BAR
    # =================================================================
    def _build_action_bar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(6, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['border'])
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel'])
        inner.pack(padx=14, pady=10, fill=tk.X)

        def _grupo(parent, titulo):
            g = tk.Frame(parent, bg=C['panel'])
            g.pack(side=tk.LEFT, padx=(0, 20))
            tk.Label(g, text=titulo.upper(), font=font(8, "bold"),
                     bg=C['panel'], fg=C['text_l']).pack(anchor='w', pady=(0, 4))
            fila = tk.Frame(g, bg=C['panel'])
            fila.pack()
            return fila

        g1 = _grupo(inner, "Archivo")
        btn_home = self._icon_button(g1, "Menú", self._volver_al_menu,
                                       C['surface'], fg=C['primary'],
                                       padx=12, pady=6)
        btn_home.pack(side=tk.LEFT, padx=2)
        ToolTip(btn_home, "Volver al menú principal de módulos.")

        self.btn_exportar = self._icon_button(
            g1, "Exportar", self.exportar_excel,
            C['primary'], padx=12, pady=6)
        self.btn_exportar.pack(side=tk.LEFT, padx=2)
        ToolTip(self.btn_exportar,
                "Exporta todos los datos e iteraciones\na un archivo Excel.")

        g2 = _grupo(inner, "Cálculo")
        tips_iter = [
            "Ejecuta la primera iteración de Newton-Raphson.",
            "Ejecuta la segunda iteración.",
            "Ejecuta la tercera iteración.",
            "Ejecuta la cuarta iteración.",
        ]
        self.btn_calcular0 = self._icon_button(
            g2, "Iter 0", self.calcular_iteracion_cero, C['pq'], padx=10, pady=6)
        self.btn_calcular0.pack(side=tk.LEFT, padx=2)
        ToolTip(self.btn_calcular0, tips_iter[0])

        self.btn_calcular1 = self._icon_button(
            g2, "Iter 1", self.calcular_iteracion_uno, C['slack'], padx=10, pady=6)
        self.btn_calcular1.pack(side=tk.LEFT, padx=2)
        ToolTip(self.btn_calcular1, tips_iter[1])

        self.btn_calcular2 = self._icon_button(
            g2, "Iter 2", self.calcular_iteracion_dos, C['pv'], padx=10, pady=6)
        self.btn_calcular2.pack(side=tk.LEFT, padx=2)
        ToolTip(self.btn_calcular2, tips_iter[2])

        self.btn_calcular3 = self._icon_button(
            g2, "Iter 3", self.calcular_iteracion_tres, C['danger'], padx=10, pady=6)
        self.btn_calcular3.pack(side=tk.LEFT, padx=2)
        ToolTip(self.btn_calcular3, tips_iter[3])

        self.btn_calcular = self._icon_button(
            g2, "Calcular todo", self._calcular_todas,
            C['primary_d'], padx=14, pady=6)
        self.btn_calcular.pack(side=tk.LEFT, padx=(10, 2))
        ToolTip(self.btn_calcular, "Ejecuta las 4 iteraciones secuencialmente.")

        g3 = _grupo(inner, "Vista")
        self.btn_diagrama = self._icon_button(
            g3, "Diagrama Unifilar", self.mostrar_diagrama_unifilar,
            C['info'], padx=12, pady=6)
        self.btn_diagrama.pack(side=tk.LEFT, padx=2)
        ToolTip(self.btn_diagrama, "Muestra el diagrama unifilar del sistema.")

        btn_chart = self._icon_button(
            g3, "Convergencia", self._show_convergence_chart,
            C['accent'], padx=12, pady=6)
        btn_chart.pack(side=tk.LEFT, padx=2)
        ToolTip(btn_chart, "Muestra la gráfica de convergencia.")

        g4 = tk.Frame(inner, bg=C['panel'])
        g4.pack(side=tk.RIGHT)
        tk.Label(g4, text="SOPORTE", font=font(8, "bold"),
                 bg=C['panel'], fg=C['text_l']).pack(anchor='e', pady=(0, 4))
        fila4 = tk.Frame(g4, bg=C['panel'])
        fila4.pack()
        btn_help = self._icon_button(fila4, "Ayuda (F1)", self._open_help,
                                       C['surface'], fg=C['accent'],
                                       padx=12, pady=6)
        btn_help.pack(side=tk.LEFT, padx=2)
        ToolTip(btn_help, "Abre el manual de uso y las fórmulas.")

    # =================================================================
    # STATUS BAR
    # =================================================================
    def _build_status_bar(self):
        C = self.C
        sb = tk.Frame(self.root, bg=C['primary_d'], height=32)
        sb.pack(fill=tk.X, side=tk.BOTTOM)
        sb.pack_propagate(False)
        tk.Frame(sb, bg=C['accent'], width=3).pack(side=tk.LEFT, fill=tk.Y)

        self.status_icon = tk.Label(sb, text="●", font=font(11),
                                      bg=C['primary_d'], fg=C['success'])
        self.status_icon.pack(side=tk.LEFT, padx=(12, 4))

        self.status_var = tk.StringVar(value="Listo.")
        tk.Label(sb, textvariable=self.status_var, font=font(9),
                 bg=C['primary_d'], fg=C['text']).pack(side=tk.LEFT)

        prog_frame = tk.Frame(sb, bg=C['primary_d'])
        prog_frame.pack(side=tk.RIGHT, padx=14)
        tk.Label(prog_frame, text="PROGRESO", font=font(7, "bold"),
                 bg=C['primary_d'], fg=C['text_l']).pack(side=tk.LEFT, padx=(0, 6))
        self.progress_canvas = tk.Canvas(prog_frame, width=120, height=8,
                                           bg=C['primary_d'], highlightthickness=0)
        self.progress_canvas.pack(side=tk.LEFT)
        self.progress_canvas.create_rectangle(0, 0, 120, 8,
                                                fill=C['border'], outline='')
        self.progress_bar = self.progress_canvas.create_rectangle(0, 0, 0, 8,
                                                                    fill=C['accent'], outline='')
        tk.Label(sb, text="v1.0  ·  Vatia", font=font_m(8),
                 bg=C['primary_d'], fg=C['text_ll']).pack(side=tk.RIGHT, padx=(0, 14))

    # =================================================================
    # HELPERS
    # =================================================================
    def _icon_button(self, parent, text, command, color, fg='white',
                     padx=14, pady=8, width=None):
        btn = tk.Button(parent, text=text, command=command,
                        bg=color, fg=fg, font=font(10, "bold"),
                        relief=tk.FLAT, cursor='hand2',
                        activebackground=self.C['accent_l'],
                        activeforeground=self.C['primary_d'],
                        padx=padx, pady=pady, borderwidth=0)
        if width: btn.config(width=width)
        return btn

    def _set_status(self, text, kind='info'):
        try:
            self.status_var.set(text)
            colors = {'info': self.C['accent'], 'ok': self.C['success'],
                      'warn': self.C['warning'], 'err': self.C['danger']}
            self.status_icon.config(fg=colors.get(kind, self.C['accent']))
        except Exception:
            pass

    def _set_progress(self, value):
        try:
            w = max(0, min(120, 120 * value))
            self.progress_canvas.coords(self.progress_bar, 0, 0, w, 8)
            if value >= 1.0:
                self.progress_canvas.itemconfig(self.progress_bar,
                                                  fill=self.C['success'])
            else:
                self.progress_canvas.itemconfig(self.progress_bar,
                                                  fill=self.C['accent'])
        except Exception:
            pass

    def _entry_valido(self, w, tipo='float'):
        try:
            v = w.get().strip()
            if v == '':
                w.configure(foreground=self.C['text']); return
            if tipo == 'float': float(v.replace(',', '.'))
            elif tipo == 'int': int(float(v.replace(',', '.')))
            w.configure(foreground=self.C['text'])
        except Exception:
            try: w.configure(foreground=self.C['danger'])
            except Exception: pass

    # =================================================================
    # ANIMACIONES
    # =================================================================
    def _update_clock(self):
        try:
            import datetime
            now = datetime.datetime.now()
            if hasattr(self, 'clock_label'):
                self.clock_label.config(text=now.strftime("%H:%M:%S"))
            if self.root.winfo_exists():
                self._clock_after_id = self.root.after(1000, self._update_clock)
        except Exception:
            pass

    def _animate_logo(self):
        if not hasattr(self, 'logo_canvas'):
            return
        try:
            if not self.logo_canvas.winfo_exists():
                return
        except Exception:
            return
        try:
            C = self.C
            self._logo_pulse = (self._logo_pulse + 0.025) % 1.0
            t = self._logo_pulse
            v_brightness = 0.35 + 0.65 * abs(math.sin(t * 2 * math.pi))
            v_color = interp_color('#0E7490', '#FFFFFF', v_brightness)
            pulse_nodes = 0.5 + 0.5 * abs(math.sin(t * 2 * math.pi))

            self.logo_canvas.delete("all")
            self._draw_vatia_logo_animated(
                cx=self._logo_cx, cy=self._logo_cy,
                size=self._logo_draw_size,
                v_color=v_color,
                pulse_nodes=pulse_nodes,
                t=t)

            self.logo_canvas.create_text(
                self._logo_cx + 55, self._logo_cy - 8,
                text="V A T I A",
                font=font(22, "bold", family=FONT_FAMILY_ALT),
                fill='#FFFFFF', anchor='w')
            self.logo_canvas.create_text(
                self._logo_cx + 57, self._logo_cy + 15,
                text="Suite Profesional de Análisis de Sistemas de Potencia",
                font=font(8, italic=True),
                fill='#CBD5E1', anchor='w')

            self._flow_spawn_counter += 1
            if self._flow_spawn_counter % 8 == 0:
                self._spawn_flow_particle()
            self._draw_flow_particles()

            self._logo_after_id = self.root.after(40, self._animate_logo)
        except tk.TclError:
            return
        except Exception:
            return

    def _draw_vatia_logo_animated(self, cx, cy, size, v_color, pulse_nodes, t):
        C = self.C
        r = size / 2
        lw_main = max(2, int(size / 14))
        lw_thin = max(1, int(size / 26))

        hex_pts = []
        for i in range(6):
            ang = 2 * math.pi * i / 6 - math.pi / 2
            hex_pts.extend([cx + r * math.cos(ang), cy + r * math.sin(ang)])
        self.logo_canvas.create_polygon(hex_pts, fill='#152741', outline='')

        tipos = [C['slack'], C['pv'], C['pq'], C['pq'], C['pv'], C['slack']]
        nodo_r = size * 0.065
        self._logo_node_positions = []
        for i in range(6):
            ang = 2 * math.pi * i / 6 - math.pi / 2
            nx = cx + r * 0.88 * math.cos(ang)
            ny = cy + r * 0.88 * math.sin(ang)
            self._logo_node_positions.append((nx, ny, i))

            self.logo_canvas.create_line(cx, cy, nx, ny,
                                           fill=C['accent'], width=lw_thin)

            phase_node = (t + i * 0.15) % 1.0
            nodo_pulse = 0.4 + 0.6 * abs(math.sin(phase_node * 2 * math.pi))
            halo_r = nodo_r * (1.2 + 0.5 * nodo_pulse)
            self.logo_canvas.create_oval(nx - halo_r, ny - halo_r,
                                           nx + halo_r, ny + halo_r,
                                           fill=C['accent'], outline='')
            self.logo_canvas.create_oval(nx - nodo_r, ny - nodo_r,
                                           nx + nodo_r, ny + nodo_r,
                                           fill=tipos[i], outline='#FFFFFF',
                                           width=max(1, int(size / 40)))

        v_h = r * 1.1
        v_w = r * 0.6
        pts = [cx - v_w, cy - v_h / 2, cx, cy + v_h / 2, cx + v_w, cy - v_h / 2]
        self.logo_canvas.create_line(*pts, fill='#0A1420',
                                       width=lw_main + 6,
                                       capstyle=tk.ROUND, joinstyle=tk.ROUND)
        self.logo_canvas.create_line(*pts, fill=v_color,
                                       width=lw_main + 2,
                                       capstyle=tk.ROUND, joinstyle=tk.ROUND)

    def _spawn_flow_particle(self):
        import random
        if not hasattr(self, '_logo_node_positions'): return
        if len(self._flow_particles) > 12: return
        idx = random.randint(0, 5)
        nx, ny, _ = self._logo_node_positions[idx]
        self._flow_particles.append({
            'x': nx, 'y': ny,
            'target_x': self._logo_cx,
            'target_y': self._logo_cy,
            'progress': 0.0,
            'speed': 0.04 + random.random() * 0.02})

    def _draw_flow_particles(self):
        activas = []
        for p in self._flow_particles:
            p['progress'] += p['speed']
            if p['progress'] >= 1.0:
                continue
            x = p['x'] + (p['target_x'] - p['x']) * p['progress']
            y = p['y'] + (p['target_y'] - p['y']) * p['progress']
            rad = 3 - 2 * p['progress']
            bright = 0.5 + 0.5 * p['progress']
            color = interp_color('#0E7490', '#FFFFFF', bright)
            self.logo_canvas.create_oval(x - rad, y - rad, x + rad, y + rad,
                                           fill=color, outline='')
            activas.append(p)
        self._flow_particles = activas

    def _animar_boton_resultados(self):
        if not hasattr(self, 'btn_resultados_finales'): return
        try:
            if not self.btn_resultados_finales.winfo_exists(): return
        except Exception:
            return
        try:
            if self.btn_resultados_finales.cget('state') == 'disabled':
                return
            self._resultados_pulse = (self._resultados_pulse + 0.08) % 1.0
            C = self.C
            bright = 0.5 + 0.5 * abs(math.sin(self._resultados_pulse * 2 * math.pi))
            color = interp_color('#F59E0B', '#FBBF24', bright)
            self.btn_resultados_finales.config(bg=color, fg='white')
            self.root.after(80, self._animar_boton_resultados)
        except Exception:
            pass

    def _deshabilitar_resultados_finales(self):
        if not hasattr(self, 'btn_resultados_finales'): return
        try:
            self.btn_resultados_finales.config(state=tk.DISABLED,
                                                 bg=self.C['surface'],
                                                 fg=self.C['text_l'])
        except Exception:
            pass

    def _detener_animaciones(self):
        """Cancela todas las animaciones pendientes."""
        for attr in ('_logo_after_id', '_sidebar_after_id', '_clock_after_id'):
            try:
                aid = getattr(self, attr, None)
                if aid:
                    self.root.after_cancel(aid)
                    setattr(self, attr, None)
            except Exception:
                pass

    def _habilitar_resultados_finales(self):
        """Habilita el botón de Resultados Finales y arranca la animación."""
        if not hasattr(self, 'btn_resultados_finales'):
            return
        try:
            self.btn_resultados_finales.config(state=tk.NORMAL,
                                                 bg='#F59E0B', fg='white')
            self._animar_boton_resultados()
        except Exception:
            pass


    # =================================================================
    # SIDEBAR UPDATE
    # =================================================================
    def _update_sidebar(self):
        if not hasattr(self, 'sidebar_labels') or not self.sidebar_labels:
            try:
                if self.root.winfo_exists():
                    self._sidebar_after_id = self.root.after(500, self._update_sidebar)
            except Exception:
                pass
            return
        try:
            n = int(self.entry_n.get())
            n_lineas = sum(1 for row in self.entries_lineas
                           if self._parse_int(row[0].get(), 0) > 0
                           and self._parse_int(row[1].get(), 0) > 0)
            self.sidebar_labels['nodos'].config(text=str(n))
            self.sidebar_labels['lineas'].config(text=str(n_lineas))

            slack = pv = pq = 0
            for cb in self.tipo_widgets:
                t = cb.get()
                if "Slack" in t: slack += 1
                elif "PV" in t: pv += 1
                else: pq += 1
            self.sidebar_labels['slack'].config(text=str(slack))
            self.sidebar_labels['pv'].config(text=str(pv))
            self.sidebar_labels['pq'].config(text=str(pq))

            if self.datos_iteracion:
                ult = max(self.datos_iteracion.keys())
                self.sidebar_labels['iter'].config(text=f"Iter {ult}")
                d = self.datos_iteracion[ult]
                dp = [abs(x) for x in d['deltaP'] if x is not None]
                dq = [abs(x) for x in d['deltaQ'] if x is not None]
                max_err = max(dp + dq) if (dp or dq) else 0.0
                if max_err < 1e-3:
                    estado, col = "CONVERGIDO", self.C['success']
                elif max_err < 0.01:
                    estado, col = "CASI", self.C['info']
                else:
                    estado, col = "PENDIENTE", self.C['warning']
                self.sidebar_labels['estado'].config(text=estado, fg=col)
                self.sidebar_labels['error'].config(text=f"{max_err:.2e}")
            else:
                self.sidebar_labels['iter'].config(text="—")
                self.sidebar_labels['estado'].config(text="SIN CALCULAR",
                                                       fg=self.C['text_l'])
                self.sidebar_labels['error'].config(text="—")

            if hasattr(self, 'status_var'):
                n_iter = len(self.datos_iteracion)
                info = f"N: {n}  ·  Líneas: {n_lineas}  ·  Iter: {n_iter}/4"
                self.status_var.set(f"● Listo    |    {info}")
        except Exception:
            pass

        try:
            if self.root.winfo_exists():
                self._sidebar_after_id = self.root.after(500, self._update_sidebar)
        except Exception:
            pass

    # =================================================================
    # GENERAR TABLAS
    # =================================================================
    def generar_tablas(self):
        try:
            n = int(self.entry_n.get())
            if n < 2:
                messagebox.showwarning("Advertencia", "Mínimo 2 nodos."); return
            if n > 20:
                messagebox.showwarning("Advertencia", "Máximo 20 nodos."); return
        except ValueError:
            messagebox.showwarning("Advertencia", "Ingrese un número entero."); return

        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)
        for lst in (self.tipo_widgets, self.V_fijo_widgets, self.theta_fijo_widgets,
                    self.Pd_widgets, self.Qd_widgets, self.Pg_widgets, self.Qg_widgets,
                    self.tipo_labels_pot, self.entries_lineas,
                    self.iter0_V_entries, self.iter0_theta_entries,
                    self.iter0_tipo_labels):
            lst.clear()
        self.datos_iteracion.clear()
        self.ventanas_resultados.clear()
        for b in self.btn_redisplay.values():
            b.config(state=tk.DISABLED)

        self.tab_nodos = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_potencias = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_lineas = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_iter0 = tk.Frame(self.notebook, bg=self.C['panel'])
        self.notebook.add(self.tab_nodos, text='   Datos Nodales   ')
        self.notebook.add(self.tab_potencias, text='   Potencias Nodales   ')
        self.notebook.add(self.tab_lineas, text='   Datos de Líneas   ')
        self.notebook.add(self.tab_iter0, text='   Iteración 0   ')

        self._crear_tab_nodos(n)
        self._crear_tab_potencias(n)
        self._crear_tab_lineas(n)
        self._crear_tab_iter0(n)

        self.tipo_widgets[0].set("Slack (3)")
        for i in range(1, n):
            self.tipo_widgets[i].set("PQ (1)")
        self._last_tipo_values = {i: self.tipo_widgets[i].get() for i in range(n)}
        self._refresh_slack_options()
        for i in range(n):
            self.actualizar_fila_iter0(i)
            self.actualizar_fila_potencias(i)

        self.btn_exportar.config(state=tk.NORMAL)
        self.btn_calcular0.config(state=tk.NORMAL)
        self.btn_calcular1.config(state=tk.DISABLED)
        self.btn_calcular2.config(state=tk.DISABLED)
        self.btn_calcular3.config(state=tk.DISABLED)
        self.btn_diagrama.config(state=tk.NORMAL)
        self._deshabilitar_resultados_finales()
        self._set_status(f"Sistema de {n} nodos configurado.", 'info')
        self._set_progress(0.0)
        self._update_sidebar()
        self.root.after(120, self._dibujar_preview)

    # =================================================================
    # TAB SCROLLEABLE
    # =================================================================
    def _crear_tab_scrolleable(self, parent):
        canvas = tk.Canvas(parent, bg=self.C['panel'], highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        interior = tk.Frame(canvas, bg=self.C['panel'])

        interior.bind('<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=interior, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        canvas.bind('<Configure>',
            lambda e: canvas.itemconfig(canvas.find_all()[0], width=e.width))

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        def _bind_wheel(e):
            canvas.bind_all('<MouseWheel>', _on_mousewheel)
        def _unbind_wheel(e):
            canvas.unbind_all('<MouseWheel>')

        canvas.bind('<Enter>', _bind_wheel)
        canvas.bind('<Leave>', _unbind_wheel)

        return interior

    # =================================================================
    # TAB DATOS NODALES
    # =================================================================
    def _crear_tab_nodos(self, n):
        C = self.C
        interior = self._crear_tab_scrolleable(self.tab_nodos)

        hdr = tk.Frame(interior, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS NODALES", font=font(12, "bold"),
                 bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Tipo de nodo  ·  Tensión  ·  Ángulo",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(interior, bg=C['panel'])
        grid.pack(padx=15, pady=15, fill=tk.X)

        headers = ["Nodo", "Tipo", "|V| (p.u.)", "θ (grados)"]
        for col_i, text in enumerate(headers):
            tk.Label(grid, text=text, font=font(9, "bold"), bg=C['head_bg'],
                     fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=col_i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(col_i, weight=1)

        for i in range(n):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']
            tk.Label(grid, text=str(i + 1), font=font(10, "bold"),
                     bg=bg, fg=C['accent'], padx=8, pady=6
                     ).grid(row=i + 1, column=0, padx=1, pady=1, sticky='nsew')
            cb = ttk.Combobox(grid, values=["Slack (3)", "PV (2)", "PQ (1)"],
                              state="readonly", font=font(9))
            cb.grid(row=i + 1, column=1, padx=1, pady=1, sticky='nsew')
            cb.set("PQ (1)")
            cb.bind("<<ComboboxSelected>>",
                    lambda e, idx=i: self._on_tipo_changed(idx))
            self.tipo_widgets.append(cb)

            for col_j in (2, 3):
                e = ttk.Entry(grid, justify='center', font=font(9))
                e.grid(row=i + 1, column=col_j, padx=1, pady=1,
                       ipady=4, sticky='nsew')
                e.bind('<KeyRelease>', lambda ev, w=e, t='float': (
                    self._entry_valido(w, t), self._dibujar_preview()))
                e.bind('<FocusOut>', lambda ev: self._dibujar_preview())
                if col_j == 2: self.V_fijo_widgets.append(e)
                else: self.theta_fijo_widgets.append(e)

    # =================================================================
    # TAB POTENCIAS NODALES
    # =================================================================
    def _crear_tab_potencias(self, n):
        C = self.C
        interior = self._crear_tab_scrolleable(self.tab_potencias)

        hdr = tk.Frame(interior, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="POTENCIAS NODALES", font=font(12, "bold"),
                 bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Potencias demandadas y generadas en MW / MVAr  ·  Base 100 MVA",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(interior, bg=C['panel'])
        grid.pack(padx=15, pady=15, fill=tk.X)

        headers = ["Nodo", "Tipo", "Pd (MW)", "Qd (MVAr)",
                   "Pg (MW)", "Qg (MVAr)"]
        for col_i, text in enumerate(headers):
            tk.Label(grid, text=text, font=font(9, "bold"), bg=C['head_bg'],
                     fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=col_i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(col_i, weight=1)

        for i in range(n):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']
            tk.Label(grid, text=str(i + 1), font=font(10, "bold"),
                     bg=bg, fg=C['accent'], padx=8, pady=6
                     ).grid(row=i + 1, column=0, padx=1, pady=1, sticky='nsew')

            lbl_tipo = tk.Label(grid, text="PQ", font=font(9, "bold"),
                                 bg=C['pq'], fg='white', padx=8, pady=6)
            lbl_tipo.grid(row=i + 1, column=1, padx=1, pady=1, sticky='nsew')
            self.tipo_labels_pot.append(lbl_tipo)

            for col_j in (2, 3, 4, 5):
                e = ttk.Entry(grid, justify='center', font=font(9))
                e.grid(row=i + 1, column=col_j, padx=1, pady=1,
                       ipady=4, sticky='nsew')
                e.insert(0, "0.0")
                e.bind('<KeyRelease>', lambda ev, w=e: (
                    self._entry_valido(w, 'float'),))
                if col_j == 2: self.Pd_widgets.append(e)
                elif col_j == 3: self.Qd_widgets.append(e)
                elif col_j == 4: self.Pg_widgets.append(e)
                else: self.Qg_widgets.append(e)

        note = tk.Frame(interior, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        note.pack(fill=tk.X, padx=15, pady=(4, 15))
        tk.Label(note,
                 text="Nota:  La potencia neta inyectada (P_esp, Q_esp) se calcula como\n"
                      "P_esp = (Pg − Pd)/S_base p.u.    y    Q_esp = (Qg − Qd)/S_base p.u.",
                 font=font_m(9), bg=C['surface'], fg=C['text_l'],
                 justify='left', anchor='w'
                 ).pack(fill=tk.X, padx=14, pady=10)

    # =================================================================
    # TAB LÍNEAS
    # =================================================================
    def _crear_tab_lineas(self, n):
        C = self.C
        interior = self._crear_tab_scrolleable(self.tab_lineas)

        hdr = tk.Frame(interior, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS DE LÍNEAS", font=font(12, "bold"),
                 bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Conecte nodos y defina admitancia serie y shunt",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(interior, bg=C['panel'])
        grid.pack(padx=15, pady=15, fill=tk.X)

        headers = ["Nodo A", "Nodo B", "G (p.u.)", "B (p.u.)", "y_shunt/2 (p.u.)"]
        for col_i, text in enumerate(headers):
            tk.Label(grid, text=text, font=font(9, "bold"), bg=C['head_bg'],
                     fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=col_i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(col_i, weight=1)

        for i in range(n):
            row = []
            for j in range(5):
                e = ttk.Entry(grid, justify='center', font=font(9))
                e.grid(row=i + 1, column=j, padx=1, pady=1,
                       ipady=4, sticky='nsew')
                tipo_val = 'int' if j in (0, 1) else 'float'
                e.bind('<KeyRelease>', lambda ev, w=e, t=tipo_val: (
                    self._entry_valido(w, t), self._dibujar_preview()))
                e.bind('<FocusOut>', lambda ev: self._dibujar_preview())
                row.append(e)
            self.entries_lineas.append(row)

    # =================================================================
    # TAB ITER 0
    # =================================================================
    def _crear_tab_iter0(self, n):
        C = self.C
        interior = self._crear_tab_scrolleable(self.tab_iter0)

        hdr = tk.Frame(interior, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="VALORES INICIALES — ITERACIÓN 0",
                 font=font(12, "bold"), bg=C['primary_d'],
                 fg='white', pady=12).pack()
        tk.Label(hdr, text="Valores de arranque para el proceso iterativo",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(interior, bg=C['panel'])
        grid.pack(padx=15, pady=15, fill=tk.X)

        for col_i, text in enumerate(["Nodo", "Tipo", "V0 (p.u.)", "θ0 (grados)"]):
            tk.Label(grid, text=text, font=font(9, "bold"), bg=C['head_bg'],
                     fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=col_i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(col_i, weight=1)

        for i in range(n):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']
            tk.Label(grid, text=str(i + 1), font=font(10, "bold"),
                     bg=bg, fg=C['accent'], padx=8, pady=6
                     ).grid(row=i + 1, column=0, padx=1, pady=1, sticky='nsew')

            tipo_str = (self.tipo_widgets[i].get()
                        if i < len(self.tipo_widgets) else "PQ (1)")
            if "Slack" in tipo_str:
                tipo_txt, tipo_bg = "Slack", C['slack']
            elif "PV" in tipo_str:
                tipo_txt, tipo_bg = "PV", C['pv']
            else:
                tipo_txt, tipo_bg = "PQ", C['pq']

            lbl = tk.Label(grid, text=tipo_txt, font=font(9, "bold"),
                           bg=tipo_bg, fg='white', padx=8, pady=6)
            lbl.grid(row=i + 1, column=1, padx=1, pady=1, sticky='nsew')

            e_v = ttk.Entry(grid, justify='center', font=font(9))
            e_v.grid(row=i + 1, column=2, padx=1, pady=1,
                     ipady=4, sticky='nsew')
            e_v.insert(0, "1.0")
            e_t = ttk.Entry(grid, justify='center', font=font(9))
            e_t.grid(row=i + 1, column=3, padx=1, pady=1,
                     ipady=4, sticky='nsew')
            e_t.insert(0, "0.0")

            self.iter0_V_entries.append(e_v)
            self.iter0_theta_entries.append(e_t)
            self.iter0_tipo_labels.append(lbl)

    # =================================================================
    # SINCRONIZACIÓN DE TIPOS
    # =================================================================
    def _on_tipo_changed(self, idx):
        if self._updating: return
        nuevo_valor = self.tipo_widgets[idx].get()
        if self._last_tipo_values.get(idx) == nuevo_valor: return
        self._last_tipo_values[idx] = nuevo_valor
        self._updating = True
        try:
            self._refresh_slack_options()
            for i in range(len(self.tipo_widgets)):
                self.actualizar_fila_iter0(i)
                self.actualizar_fila_potencias(i)
            self._dibujar_preview()
            self._update_sidebar()
            self._set_status(f"Tipo del nodo {idx + 1} actualizado a {nuevo_valor}.")
        finally:
            self._updating = False

    def _refresh_slack_options(self):
        slack_idx = None
        for i, cb in enumerate(self.tipo_widgets):
            if "Slack" in cb.get():
                slack_idx = i; break
        for i, cb in enumerate(self.tipo_widgets):
            current_val = cb.get()
            if i == slack_idx:
                new_vals = ["Slack (3)", "PV (2)", "PQ (1)"]
            elif slack_idx is not None:
                new_vals = ["PV (2)", "PQ (1)"]
            else:
                new_vals = ["Slack (3)", "PV (2)", "PQ (1)"]
            old_vals = list(cb['values']) if cb['values'] else []
            if old_vals != new_vals:
                cb.config(values=new_vals)
                if current_val in new_vals:
                    cb.set(current_val)
                elif (current_val == "Slack (3)" and
                      slack_idx is not None and i != slack_idx):
                    cb.set("PQ (1)")
                    self._last_tipo_values[i] = "PQ (1)"

    def _sincronizar_tipos(self):
        try:
            if not hasattr(self, 'tipo_widgets'): return
            for i in range(len(self.tipo_widgets)):
                self.actualizar_fila_iter0(i)
                self.actualizar_fila_potencias(i)
        except Exception:
            pass

    def actualizar_fila_iter0(self, idx):
        if idx >= len(self.iter0_V_entries):
            return
        if idx >= len(self.iter0_tipo_labels):
            return
        C = self.C
        tipo_str = self.tipo_widgets[idx].get()
        lbl = self.iter0_tipo_labels[idx]
        if "Slack" in tipo_str:
            lbl.config(text="Slack", bg=C['slack'])
        elif "PV" in tipo_str:
            lbl.config(text="PV", bg=C['pv'])
        else:
            lbl.config(text="PQ", bg=C['pq'])

    def actualizar_fila_potencias(self, idx):
        if idx >= len(self.tipo_labels_pot):
            return
        C = self.C
        tipo_str = self.tipo_widgets[idx].get()
        lbl = self.tipo_labels_pot[idx]
        if "Slack" in tipo_str:
            lbl.config(text="Slack", bg=C['slack'])
        elif "PV" in tipo_str:
            lbl.config(text="PV", bg=C['pv'])
        else:
            lbl.config(text="PQ", bg=C['pq'])

    # =================================================================
    # PREVIEW
    # =================================================================
    def _dibujar_preview(self):
        if self._updating: return
        canvas = getattr(self, 'preview_canvas', None)
        if canvas is None: return
        try:
            n = int(self.entry_n.get())
        except Exception:
            return
        if n < 2 or len(self.tipo_widgets) != n: return

        canvas.delete("all")
        W = canvas.winfo_width(); H = canvas.winfo_height()
        if W < 50 or H < 50: W, H = 520, 520
        C = self.C

        canvas.create_rectangle(0, 0, W, H, fill='#FFFFFF', outline='')

        canvas.create_text(W / 2, 26, text="Diagrama Unifilar — Vista Previa",
                           font=font(13, "bold"), fill=C['primary'])
        canvas.create_line(W / 2 - 100, 42, W / 2 + 100, 42,
                           fill=C['accent'], width=2)

        tipos = []
        for cb in self.tipo_widgets:
            s = cb.get()
            tipos.append(3 if "Slack" in s else (2 if "PV" in s else 1))

        V_complex_calc = None
        if self.datos_iteracion and 0 in self.datos_iteracion:
            try:
                d0 = self.datos_iteracion[0]
                if d0.get('n') == n:
                    V_complex_calc = [d0['V'][i] * complex(math.cos(d0['theta'][i]),
                                                            math.sin(d0['theta'][i]))
                                      for i in range(n)]
            except Exception:
                V_complex_calc = None

        lineas = []
        for row in self.entries_lineas:
            a = self._parse_int(row[0].get(), 0); b = self._parse_int(row[1].get(), 0)
            G = self._parse_float(row[2].get(), 0.0); B = self._parse_float(row[3].get(), 0.0)
            ysh = self._parse_float(row[4].get(), 0.0)
            if a > 0 and b > 0 and a <= n and b <= n and a != b:
                lineas.append((a - 1, b - 1, G, B, ysh))

        if not lineas:
            canvas.create_text(
                W / 2, H / 2,
                text="Sin líneas definidas\n\n"
                     "Complete la pestaña 'Datos de Líneas'\n"
                     "para visualizar el diagrama unifilar",
                font=font(11), fill=C['text_ll'], justify=tk.CENTER)
            return

        cx, cy = W / 2, H / 2 + 20
        R = min(W, H) * 0.30
        pos = []
        for i in range(n):
            ang = 2 * math.pi * i / n - math.pi / 2
            pos.append((cx + R * math.cos(ang), cy + R * math.sin(ang)))

        for (a, b, G, B, ysh) in lineas:
            x1, y1 = pos[a]; x2, y2 = pos[b]
            canvas.create_line(x1 + 2, y1 + 2, x2 + 2, y2 + 2, width=4,
                               fill='#0A0E15', capstyle=tk.ROUND)
            canvas.create_line(x1, y1, x2, y2, width=3,
                               fill=C['gray_dark'], capstyle=tk.ROUND)
            P_ab = None
            if V_complex_calc is not None:
                try:
                    Va = V_complex_calc[a]; Vb = V_complex_calc[b]
                    I_ab = (Va - Vb) * complex(G, B) + Va * complex(0, ysh)
                    P_ab = (Va * I_ab.conjugate()).real
                except Exception:
                    P_ab = None
            if P_ab is not None:
                fwd = P_ab >= 0
                t_pos = 0.20
                ax0 = x1 + (x2 - x1) * t_pos; ay0 = y1 + (y2 - y1) * t_pos
                adx = (x2 - x1) if fwd else (x1 - x2)
                ady = (y2 - y1) if fwd else (y1 - y2)
                alen = math.hypot(adx, ady) or 1.0
                aux, auy = adx / alen, ady / alen
                px, py = -auy, aux
                sz, hw = 12, 7
                canvas.create_polygon(ax0 + aux * sz / 2, ay0 + auy * sz / 2,
                                       ax0 - aux * sz / 2 + px * hw,
                                       ay0 - auy * sz / 2 + py * hw,
                                       ax0 - aux * sz / 2 - px * hw,
                                       ay0 - auy * sz / 2 - py * hw,
                                       fill=C['accent'], outline='white', width=1)
                        # --- Etiqueta de línea desplazada perpendicularmente ---
            lbl = f"Y={G:+.3f}{'+' if B >= 0 else '-'}j{abs(B):.3f}"
            if P_ab is not None:
                lbl += f"\nP={P_ab:+.3f}"

            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            # Vector perpendicular normalizado
            L = math.hypot(x2 - x1, y2 - y1) or 1.0
            perp_x = -(y2 - y1) / L
            perp_y = (x2 - x1) / L
            # Offset alternado según el índice de la línea
            offset = 22 if (len(lineas) > 1 and
                            (lineas.index((a, b, G, B, ysh)) % 2 == 0)) else -22
            lx_lbl = mx + perp_x * offset
            ly_lbl = my + perp_y * offset

            tid = canvas.create_text(lx_lbl, ly_lbl, text=lbl,
                                     font=font_m(7, "bold"),
                                     fill=C['accent'], justify=tk.CENTER)
            bb = canvas.bbox(tid)
            if bb:
                canvas.create_rectangle(bb[0] - 4, bb[1] - 2,
                                         bb[2] + 4, bb[3] + 2,
                                         fill='#FFFFFF', outline=C['accent'],
                                         width=1)
                canvas.tag_raise(tid)
            # Línea guía punteada del centro a la etiqueta
            canvas.create_line(mx, my, lx_lbl, ly_lbl,
                               fill=C['border'], width=1, dash=(2, 2))

        color_map = {1: C['pq'], 2: C['pv'], 3: C['slack']}
        ring_map = {1: '#FFEDD5', 2: '#EDE9FE', 3: '#D1FAE5'}
        name_map = {1: 'PQ', 2: 'PV', 3: 'SLK'}
        for i, (x, y) in enumerate(pos):
            tipo = tipos[i]
            col = color_map[tipo]; ring = ring_map[tipo]
            canvas.create_line(x - 24, y, x + 24, y, width=10,
                               fill=C['gray_dark'], capstyle=tk.ROUND)
            canvas.create_oval(x - 17, y - 17, x + 17, y + 17, fill=ring, outline='')
            canvas.create_oval(x - 13, y - 13, x + 13, y + 13, fill=col,
                               outline='white', width=2)
            canvas.create_text(x, y, text=str(i + 1), fill='white',
                               font=font(10, "bold"))
            canvas.create_text(x, y + 30, text=name_map[tipo], font=font(8, "bold"),
                               fill=col)

        lx, ly = 12, H - 82
        canvas.create_rectangle(lx + 2, ly + 2, lx + 158, ly + 70,
                                fill='#0A0E15', outline='')
        canvas.create_rectangle(lx, ly, lx + 156, ly + 68, fill=C['panel'],
                                outline=C['border'], width=1)
        canvas.create_text(lx + 78, ly + 14, text="LEYENDA",
                           font=font(8, "bold"), fill=C['accent'])
        canvas.create_line(lx + 40, ly + 24, lx + 116, ly + 24, fill=C['border'])
        items = [(C['slack'], 'Slack'), (C['pv'], 'PV'), (C['pq'], 'PQ')]
        for k, (col, txt) in enumerate(items):
            yk = ly + 32 + k * 12
            canvas.create_oval(lx + 14, yk - 5, lx + 24, yk + 5, fill=col,
                               outline='white', width=1)
            canvas.create_text(lx + 32, yk, text=txt, font=font(8),
                               fill=C['text'], anchor='w')
        canvas.create_text(W - 14, H - 12, text=f"{n} nodos · {len(lineas)} línea(s)",
                           font=font(8, italic=True), fill=C['text_l'], anchor='se')

    # =================================================================
    # HELPERS DE PRESENTACIÓN
    # =================================================================
    def _crear_tarjeta_kpi(self, parent, titulo, valor, subtitulo, color):
        card = tk.Frame(parent, bg=self.C['panel'], highlightthickness=1,
                        highlightbackground=self.C['border'])
        card.pack(side=tk.LEFT, padx=6, pady=2, fill=tk.Y)
        tk.Frame(card, bg=color, width=5).pack(side=tk.LEFT, fill=tk.Y)
        body = tk.Frame(card, bg=self.C['panel'])
        body.pack(side=tk.LEFT, padx=14, pady=10)
        tk.Label(body, text=titulo.upper(), font=font(8, "bold"),
                 bg=self.C['panel'], fg=self.C['text_l']).pack(anchor='w')
        tk.Label(body, text=valor, font=font(16, "bold"),
                 bg=self.C['panel'], fg=color).pack(anchor='w', pady=(2, 0))
        tk.Label(body, text=subtitulo, font=font(8, italic=True),
                 bg=self.C['panel'], fg=self.C['text_l']).pack(anchor='w')

    def _crear_seccion(self, parent, titulo, color):
        hdr = tk.Frame(parent, bg=self.C['surface'], highlightthickness=1,
                       highlightbackground=self.C['border'])
        hdr.pack(fill=tk.X, pady=(6, 10))
        tk.Frame(hdr, bg=color, width=5).pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(hdr, text=titulo, font=font(11, "bold"),
                 bg=self.C['surface'], fg=self.C['text'],
                 pady=8).pack(side=tk.LEFT, padx=14)

    def _crear_tabla_datos(self, parent, headers, rows, col_widths=None,
                            colorear_celdas=None):
        C = self.C
        container = tk.Frame(parent, bg=C['border'])
        container.pack(fill=tk.X, pady=(0, 10))

        head = tk.Frame(container, bg=C['primary_d'])
        head.pack(fill=tk.X)
        for i, h in enumerate(headers):
            head.grid_columnconfigure(i, weight=1, uniform='col')
            tk.Label(head, text=h, font=font(9, "bold"),
                     bg=C['primary_d'], fg='white',
                     padx=10, pady=8
                     ).grid(row=0, column=i, sticky='nsew', padx=1, pady=1)

        body = tk.Frame(container, bg=C['border'])
        body.pack(fill=tk.X)
        for r_idx, row_data in enumerate(rows):
            bg = C['row_even'] if r_idx % 2 == 0 else C['row_odd']
            for c_idx, val in enumerate(row_data):
                body.grid_columnconfigure(c_idx, weight=1, uniform='col')
                cell_bg = bg; cell_fg = C['text']
                if colorear_celdas:
                    ov = colorear_celdas(r_idx, c_idx, val)
                    if ov:
                        if isinstance(ov, tuple):
                            cell_bg, cell_fg = ov
                        else:
                            cell_bg = ov
                tk.Label(body, text=str(val), font=font_m(9),
                         bg=cell_bg, fg=cell_fg,
                         padx=10, pady=6
                         ).grid(row=r_idx, column=c_idx,
                                sticky='nsew', padx=1, pady=1)

        return container

    def _crear_matriz_bonita(self, parent, matriz, highlight_diag=False,
                              cell_h=28, fuente=None):
        C = self.C
        if fuente is None:
            fuente = font_m(9)
        n_rows = len(matriz); n_cols = len(matriz[0]) if matriz else 0
        if n_rows == 0 or n_cols == 0:
            return tk.Frame(parent, bg=C['panel'], width=1, height=1)
        max_len = max(len(str(matriz[i][j]))
                      for i in range(n_rows) for j in range(n_cols))
        cell_w = max(72, max_len * 7 + 22)
        bracket_w = 16
        W = bracket_w * 2 + n_cols * cell_w
        H = n_rows * cell_h
        canvas = tk.Canvas(parent, width=W, height=H, bg=C['panel'],
                           highlightthickness=0)
        for i in range(n_rows):
            for j in range(n_cols):
                x0 = bracket_w + j * cell_w; y0 = i * cell_h
                bg = C['row_even'] if i % 2 == 0 else C['row_odd']
                fg = C['text']
                if highlight_diag and i == j:
                    bg = '#DBEAFE'; fg = C['primary_d']
                canvas.create_rectangle(x0, y0, x0 + cell_w, y0 + cell_h,
                                         fill=bg, outline='')
                canvas.create_text(x0 + cell_w - 10, y0 + cell_h / 2,
                                    text=str(matriz[i][j]),
                                    font=fuente, fill=fg, anchor='e')
        canvas.create_line(bracket_w - 4, 2, 5, 2, fill=C['accent'], width=3)
        canvas.create_line(5, 2, 5, H - 2, fill=C['accent'], width=3)
        canvas.create_line(5, H - 2, bracket_w - 4, H - 2, fill=C['accent'], width=3)
        xr = bracket_w + n_cols * cell_w
        canvas.create_line(xr + 4, 2, xr + bracket_w - 4, 2, fill=C['accent'], width=3)
        canvas.create_line(xr + bracket_w - 4, 2, xr + bracket_w - 4, H - 2,
                           fill=C['accent'], width=3)
        canvas.create_line(xr + bracket_w - 4, H - 2, xr + 4, H - 2,
                           fill=C['accent'], width=3)
        return canvas

    def _crear_vector_columna(self, parent, valores, titulo=None):
        C = self.C
        wrap = tk.Frame(parent, bg=C['panel'])
        if titulo:
            tk.Label(wrap, text=titulo, font=font(9, "bold"),
                     bg=C['panel'], fg=C['primary']).pack(pady=(0, 4))
        self._crear_matriz_bonita(wrap, [[v] for v in valores]).pack()
        return wrap

    # =================================================================
    # CÁLCULOS
    # =================================================================
    def calcular_iteracion_cero(self):
        try:
            self._deshabilitar_resultados_finales()
            n = int(self.entry_n.get())
            nodos = []
            for i in range(n):
                tipo_str = self.tipo_widgets[i].get()
                tipo = 3 if "Slack" in tipo_str else (2 if "PV" in tipo_str else 1)

                Pd = self._parse_float(self.Pd_widgets[i].get(), 0.0)
                Qd = self._parse_float(self.Qd_widgets[i].get(), 0.0)
                Pg = self._parse_float(self.Pg_widgets[i].get(), 0.0)
                Qg = self._parse_float(self.Qg_widgets[i].get(), 0.0)

                Pesp = (Pg - Pd) / self.BASE_MVA
                Qesp = (Qg - Qd) / self.BASE_MVA

                if tipo in [2, 3]:
                    V = self._parse_float(self.V_fijo_widgets[i].get(), 1.0)
                    theta = self._parse_float(self.theta_fijo_widgets[i].get(), 0.0)
                else:
                    V = None; theta = None

                nodos.append({
                    'tipo': tipo, 'Pesp': Pesp, 'Qesp': Qesp,
                    'V_fijo': V, 'theta_fijo': theta,
                    'Pd_MW': Pd, 'Qd_MVAr': Qd,
                    'Pg_MW': Pg, 'Qg_MVAr': Qg,
                })

            lineas = []
            for row in self.entries_lineas:
                a = self._parse_int(row[0].get(), 0); b = self._parse_int(row[1].get(), 0)
                G = self._parse_float(row[2].get(), 0.0); B = self._parse_float(row[3].get(), 0.0)
                y_shunt = self._parse_float(row[4].get(), 0.0)
                if a == 0 or b == 0: continue
                if a < 1 or a > n or b < 1 or b > n:
                    messagebox.showwarning("Nodo inválido",
                                           f"Línea {a}-{b} ignorada.")
                    continue
                lineas.append({'a': a - 1, 'b': b - 1,
                               'G': G, 'B': B, 'y_shunt': y_shunt})

            V0 = []; theta0 = []
            for i in range(n):
                v = self._parse_float(self.iter0_V_entries[i].get(), 1.0)
                t = self._parse_float(self.iter0_theta_entries[i].get(), 0.0)
                V0.append(v); theta0.append(math.radians(t))
            for i in range(n):
                if nodos[i]['tipo'] in [2, 3]:
                    V0[i] = nodos[i]['V_fijo']
                    theta0[i] = math.radians(nodos[i]['theta_fijo'])

            Ybus = [[0j] * n for _ in range(n)]
            for line in lineas:
                a, b = line['a'], line['b']
                G = line['G']; B = line['B']; y_shunt = line['y_shunt']
                y_serie = complex(G, B)
                Ybus[a][a] += y_serie + complex(0, y_shunt)
                Ybus[b][b] += y_serie + complex(0, y_shunt)
                Ybus[a][b] -= y_serie; Ybus[b][a] -= y_serie

            V = [V0[i] * complex(math.cos(theta0[i]), math.sin(theta0[i]))
                 for i in range(n)]
            I = [0j for _ in range(n)]
            for i in range(n):
                for j in range(n):
                    I[i] += Ybus[i][j] * V[j]

            S_calc = [V[i] * complex(I[i].real, -I[i].imag) for i in range(n)]
            P_calc = [S.real for S in S_calc]; Q_calc = [S.imag for S in S_calc]

            deltaP = [None] * n; deltaQ = [None] * n
            for i in range(n):
                if nodos[i]['tipo'] != 3:
                    deltaP[i] = nodos[i]['Pesp'] - P_calc[i]
                if nodos[i]['tipo'] == 1:
                    deltaQ[i] = nodos[i]['Qesp'] - Q_calc[i]

            Ybus_np = np.array([[Ybus[i][j] for j in range(n)] for i in range(n)],
                               dtype=np.complex128)
            G_np = np.real(Ybus_np); B_np = np.imag(Ybus_np)
            idx_pq = [i for i in range(n) if nodos[i]['tipo'] == 1]

            J0, theta_idx, V_idx = self.construir_jacobiano(
                n, nodos, G_np, B_np, theta0, V0, P_calc, Q_calc, idx_pq)

            mismatch_vec = []
            for i in range(n):
                if nodos[i]['tipo'] != 3:
                    mismatch_vec.append(deltaP[i] if deltaP[i] is not None else 0.0)
            for i in idx_pq:
                mismatch_vec.append(deltaQ[i] if deltaQ[i] is not None else 0.0)
            mismatch_vec = np.array(mismatch_vec, dtype=np.float64)

            try:
                delta_x = np.linalg.solve(J0, mismatch_vec)
            except np.linalg.LinAlgError as e:
                messagebox.showerror("Error", f"Jacobiano singular: {e}")
                return

            delta_theta0 = np.zeros(n); delta_V0 = np.zeros(n)
            for i in range(n):
                if nodos[i]['tipo'] != 3:
                    delta_theta0[i] = delta_x[theta_idx[i]]
            for i in idx_pq:
                delta_V0[i] = delta_x[V_idx[i]]

            datos_iter0 = {
                'n': n, 'nodos': nodos, 'Ybus': Ybus,
                'V': V0.copy(), 'theta': theta0.copy(),
                'P_calc': P_calc.copy(), 'Q_calc': Q_calc.copy(),
                'deltaP': deltaP.copy(), 'deltaQ': deltaQ.copy(),
                'G': G_np, 'B': B_np, 'idx_pq': idx_pq,
                'J': J0, 'theta_idx': theta_idx, 'V_idx': V_idx,
                'delta_theta': delta_theta0, 'delta_V': delta_V0,
                'V_prev': V0.copy(), 'theta_prev': theta0.copy(),
                'deltaP_prev': deltaP.copy(), 'deltaQ_prev': deltaQ.copy(),
                'P_calc_prev': P_calc.copy(), 'Q_calc_prev': Q_calc.copy(),
                'mismatch': mismatch_vec, 'delta_x': delta_x}

            self.datos_iteracion[0] = datos_iter0
            self.mostrar_resultados_iterN(0, datos_iter0)
            self.btn_redisplay0.config(state=tk.NORMAL)
            self.btn_calcular0.config(state=tk.DISABLED)
            self.btn_calcular1.config(state=tk.NORMAL)
            self.btn_calcular2.config(state=tk.DISABLED)
            self.btn_calcular3.config(state=tk.DISABLED)
            self._set_progress(0.25)
            self._update_sidebar()
            self._set_status("Iteración 0 calculada correctamente.", 'ok')
        except Exception as e:
            messagebox.showerror("Error", f"Error:\n{str(e)}")

    def construir_jacobiano(self, n, nodos, G, B, theta, V, P_calc, Q_calc, idx_pq):
        theta_idx = {}; var_idx = 0
        for i in range(n):
            if nodos[i]['tipo'] != 3:
                theta_idx[i] = var_idx; var_idx += 1
        V_idx = {}
        for i in idx_pq:
            V_idx[i] = var_idx; var_idx += 1
        m = var_idx; J = np.zeros((m, m))
        row = 0
        for i in range(n):
            if nodos[i]['tipo'] == 3: continue
            for j in range(n):
                if nodos[j]['tipo'] == 3: continue
                if i == j:
                    J[row, theta_idx[j]] = -Q_calc[i] - B[i][i] * V[i] ** 2
                else:
                    J[row, theta_idx[j]] = V[i] * V[j] * (
                        G[i][j] * math.sin(theta[i] - theta[j])
                        - B[i][j] * math.cos(theta[i] - theta[j]))
            for k in idx_pq:
                if i == k:
                    J[row, V_idx[k]] = (0.0 if abs(V[i]) < 1e-12
                                         else P_calc[i]/V[i] + G[i][i]*V[i])
                else:
                    J[row, V_idx[k]] = V[i] * (
                        G[i][k]*math.cos(theta[i]-theta[k])
                        + B[i][k]*math.sin(theta[i]-theta[k]))
            row += 1
        for i in idx_pq:
            for j in range(n):
                if nodos[j]['tipo'] == 3: continue
                if i == j:
                    J[row, theta_idx[j]] = P_calc[i] - G[i][i] * V[i] ** 2
                else:
                    J[row, theta_idx[j]] = -V[i]*V[j]*(
                        G[i][j]*math.cos(theta[i]-theta[j])
                        + B[i][j]*math.sin(theta[i]-theta[j]))
            for k in idx_pq:
                if i == k:
                    J[row, V_idx[k]] = (0.0 if abs(V[i]) < 1e-12
                                         else Q_calc[i]/V[i] - B[i][i]*V[i])
                else:
                    J[row, V_idx[k]] = V[i] * (
                        G[i][k]*math.sin(theta[i]-theta[k])
                        - B[i][k]*math.cos(theta[i]-theta[k]))
            row += 1
        return J, theta_idx, V_idx

    def calcular_siguiente_iteracion(self, datos_prev):
        n = datos_prev['n']; nodos = datos_prev['nodos']
        Ybus = datos_prev['Ybus']
        V_prev = datos_prev['V'].copy(); theta_prev = datos_prev['theta'].copy()
        P_calc_prev = datos_prev['P_calc']; Q_calc_prev = datos_prev['Q_calc']
        deltaP_prev = datos_prev['deltaP']; deltaQ_prev = datos_prev['deltaQ']
        G = datos_prev['G']; B = datos_prev['B']; idx_pq = datos_prev['idx_pq']

        J_usado, theta_idx, V_idx = self.construir_jacobiano(
            n, nodos, G, B, theta_prev, V_prev,
            P_calc_prev, Q_calc_prev, idx_pq)

        mismatch_usado = []
        for i in range(n):
            if nodos[i]['tipo'] != 3:
                mismatch_usado.append(deltaP_prev[i]
                                      if deltaP_prev[i] is not None else 0.0)
        for i in idx_pq:
            mismatch_usado.append(deltaQ_prev[i]
                                  if deltaQ_prev[i] is not None else 0.0)
        mismatch_usado = np.array(mismatch_usado, dtype=np.float64)

        if np.linalg.cond(J_usado) > 1e12:
            messagebox.showerror("Error", "Jacobiano casi singular.")
            return None

        try:
            delta_x = np.linalg.solve(J_usado, mismatch_usado)
        except np.linalg.LinAlgError as e:
            messagebox.showerror("Error", f"Jacobiano singular: {e}")
            return None

        delta_theta_usado = np.zeros(n); delta_V_usado = np.zeros(n)
        for i in range(n):
            if nodos[i]['tipo'] != 3:
                delta_theta_usado[i] = delta_x[theta_idx[i]]
        for i in idx_pq:
            delta_V_usado[i] = delta_x[V_idx[i]]

        theta_new = theta_prev + delta_theta_usado
        V_new = V_prev + delta_V_usado

        V_complex = [V_new[i]*complex(math.cos(theta_new[i]),
                                        math.sin(theta_new[i]))
                     for i in range(n)]
        I = [0j for _ in range(n)]
        for i in range(n):
            for j in range(n):
                I[i] += Ybus[i][j] * V_complex[j]
        S_calc = [V_complex[i]*complex(I[i].real, -I[i].imag)
                  for i in range(n)]
        P_calc_new = [S.real for S in S_calc]
        Q_calc_new = [S.imag for S in S_calc]

        deltaP_new = [None]*n; deltaQ_new = [None]*n
        for i in range(n):
            if nodos[i]['tipo'] != 3:
                deltaP_new[i] = nodos[i]['Pesp'] - P_calc_new[i]
            if nodos[i]['tipo'] == 1:
                deltaQ_new[i] = nodos[i]['Qesp'] - Q_calc_new[i]

        J_nuevo, _, _ = self.construir_jacobiano(
            n, nodos, G, B, theta_new, V_new,
            P_calc_new, Q_calc_new, idx_pq)

        mismatch_nuevo = []
        for i in range(n):
            if nodos[i]['tipo'] != 3:
                mismatch_nuevo.append(deltaP_new[i]
                                      if deltaP_new[i] is not None else 0.0)
        for i in idx_pq:
            mismatch_nuevo.append(deltaQ_new[i]
                                  if deltaQ_new[i] is not None else 0.0)
        mismatch_nuevo = np.array(mismatch_nuevo, dtype=np.float64)

        try:
            delta_x_nuevo = np.linalg.solve(J_nuevo, mismatch_nuevo)
        except np.linalg.LinAlgError:
            delta_x_nuevo = np.zeros(len(mismatch_nuevo))

        delta_theta_nuevo = np.zeros(n); delta_V_nuevo = np.zeros(n)
        for i in range(n):
            if nodos[i]['tipo'] != 3:
                delta_theta_nuevo[i] = delta_x_nuevo[theta_idx[i]]
        for i in idx_pq:
            delta_V_nuevo[i] = delta_x_nuevo[V_idx[i]]

        return {'n': n, 'nodos': nodos, 'Ybus': Ybus,
            'V': V_new.copy(), 'theta': theta_new.copy(),
            'P_calc': P_calc_new.copy(), 'Q_calc': Q_calc_new.copy(),
            'deltaP': deltaP_new.copy(), 'deltaQ': deltaQ_new.copy(),
            'G': G, 'B': B, 'idx_pq': idx_pq,
            'J': J_nuevo, 'J_usado': J_usado,
            'theta_idx': theta_idx, 'V_idx': V_idx,
            'delta_theta_usado': delta_theta_usado.copy(),
            'delta_V_usado': delta_V_usado.copy(),
            'delta_theta': delta_theta_nuevo.copy(),
            'delta_V': delta_V_nuevo.copy(),
            'V_prev': V_prev.copy(), 'theta_prev': theta_prev.copy(),
            'deltaP_prev': deltaP_prev.copy(), 'deltaQ_prev': deltaQ_prev.copy(),
            'P_calc_prev': P_calc_prev.copy(), 'Q_calc_prev': Q_calc_prev.copy(),
            'mismatch_usado': mismatch_usado.copy(),
            'mismatch_actual': mismatch_nuevo.copy(),
            'delta_x': delta_x_nuevo}

    def calcular_iteracion_uno(self):
        if 0 not in self.datos_iteracion:
            messagebox.showwarning("Advertencia", "Primero calcula la iteración 0.")
            return
        dn = self.calcular_siguiente_iteracion(self.datos_iteracion[0])
        if dn is None: return
        self.datos_iteracion[1] = dn
        self.mostrar_resultados_iterN(1, dn)
        self.btn_redisplay1.config(state=tk.NORMAL)
        self.btn_calcular1.config(state=tk.DISABLED)
        self.btn_calcular2.config(state=tk.NORMAL)
        self.btn_calcular3.config(state=tk.DISABLED)
        self._set_progress(0.50); self._update_sidebar()
        self._set_status("Iteración 1 calculada.", 'ok')

    def calcular_iteracion_dos(self):
        if 1 not in self.datos_iteracion:
            messagebox.showwarning("Advertencia", "Primero calcula la iteración 1.")
            return
        dn = self.calcular_siguiente_iteracion(self.datos_iteracion[1])
        if dn is None: return
        self.datos_iteracion[2] = dn
        self.mostrar_resultados_iterN(2, dn)
        self.btn_redisplay2.config(state=tk.NORMAL)
        self.btn_calcular2.config(state=tk.DISABLED)
        self.btn_calcular3.config(state=tk.NORMAL)
        self._set_progress(0.75); self._update_sidebar()
        self._set_status("Iteración 2 calculada.", 'ok')

    def calcular_iteracion_tres(self):
        if 2 not in self.datos_iteracion:
            messagebox.showwarning("Advertencia", "Primero calcula la iteración 2.")
            return
        dn = self.calcular_siguiente_iteracion(self.datos_iteracion[2])
        if dn is None: return
        self.datos_iteracion[3] = dn
        self.mostrar_resultados_iterN(3, dn)
        self.btn_redisplay3.config(state=tk.NORMAL)
        self.btn_calcular3.config(state=tk.DISABLED)
        self._set_progress(1.0); self._update_sidebar()
        self._set_status("Iteración 3 calculada. Proceso completado.", 'ok')
        self._habilitar_resultados_finales()

    def _calcular_todas(self):
        try:
            if not self.tipo_widgets:
                messagebox.showwarning("Sin datos",
                    "Primero configure el sistema con 'Generar Tablas'.")
                return
            self.datos_iteracion.clear()
            for b in self.btn_redisplay.values():
                b.config(state=tk.DISABLED)
            self._deshabilitar_resultados_finales()
            self._set_progress(0.0)
            self._set_status("Iniciando cálculo secuencial...", 'info')
            self.root.update_idletasks()

            self.calcular_iteracion_cero()
            if 0 not in self.datos_iteracion: return
            self.root.update_idletasks()

            self.calcular_iteracion_uno()
            if 1 not in self.datos_iteracion: return
            self.root.update_idletasks()

            self.calcular_iteracion_dos()
            if 2 not in self.datos_iteracion: return
            self.root.update_idletasks()

            self.calcular_iteracion_tres()
            if 3 in self.datos_iteracion:
                self._set_progress(1.0)
                self._set_status("Cálculo completado. 4 iteraciones ejecutadas.", 'ok')
                self.btn_calcular.config(state=tk.DISABLED)
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Error en el cálculo:\n{e}")

    # =================================================================
    # RESULTADOS POR ITERACIÓN
    # =================================================================
    def mostrar_resultados_iterN(self, num_iter, datos):
        if num_iter in self.ventanas_resultados and self.ventanas_resultados[num_iter]:
            try:
                self.ventanas_resultados[num_iter].lift(); return
            except Exception:
                pass
        C = self.C
        n = datos['n']; nodos = datos['nodos']; idx_pq = datos['idx_pq']
        Ybus = datos['Ybus']; V_actual = datos['V']; theta_actual = datos['theta']
        P_calc = datos['P_calc']; Q_calc = datos['Q_calc']
        deltaP = datos['deltaP']; deltaQ = datos['deltaQ']

        ventana = tk.Toplevel(self.root)
        ventana.title(f"Resultados — Iteración {num_iter}")
        ventana.geometry("1440x920"); ventana.minsize(1150, 720)
        ventana.configure(bg=C['bg'])
        self.ventanas_resultados[num_iter] = ventana

        hdr = tk.Frame(ventana, bg=C['primary_d'], height=70)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['accent'], height=4).pack(fill=tk.X, side=tk.TOP)

        hl = tk.Frame(hdr, bg=C['primary_d'])
        hl.pack(side=tk.LEFT, padx=24, pady=12)
        badge = tk.Frame(hl, bg=C['accent'], width=54, height=54)
        badge.pack(side=tk.LEFT, padx=(0, 16)); badge.pack_propagate(False)
        tk.Label(badge, text=str(num_iter), font=font(22, "bold"),
                 bg=C['accent'], fg=C['primary_d']).pack(expand=True)
        ht = tk.Frame(hl, bg=C['primary_d']); ht.pack(side=tk.LEFT)
        tk.Label(ht, text=f"Resultados — Iteración {num_iter}",
                 font=font(16, "bold"), bg=C['primary_d'], fg='white'
                 ).pack(anchor='w')
        tk.Label(ht, text="Análisis detallado del flujo de carga Newton-Raphson",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(anchor='w')

        kpi_outer = tk.Frame(ventana, bg=C['bg'])
        kpi_outer.pack(fill=tk.X, padx=14, pady=12)
        kpi_row = tk.Frame(kpi_outer, bg=C['bg']); kpi_row.pack(fill=tk.X)

        dp_vals = [abs(deltaP[i]) for i in range(n) if deltaP[i] is not None]
        dq_vals = [abs(deltaQ[i]) for i in range(n) if deltaQ[i] is not None]
        max_dp = max(dp_vals) if dp_vals else 0.0
        max_dq = max(dq_vals) if dq_vals else 0.0
        max_err = max(max_dp, max_dq)
        if max_err < 1e-3: estado, ec = "CONVERGIDO", C['success']
        elif max_err < 0.01: estado, ec = "CASI CONVERGE", C['info']
        elif max_err < 0.1: estado, ec = "CONVERGIENDO", C['warning']
        else: estado, ec = "ALTO ERROR", C['danger']

        self._crear_tarjeta_kpi(kpi_row, "Error máx |ΔP|", f"{max_dp:.2e} p.u.",
                                "Peor desbalance de P activa", C['accent'])
        self._crear_tarjeta_kpi(kpi_row, "Error máx |ΔQ|", f"{max_dq:.2e} p.u.",
                                "Peor desbalance de Q reactiva", C['pv'])
        self._crear_tarjeta_kpi(kpi_row, "Estado", estado,
                                "Criterio tol. 1e-3 p.u.", ec)
        self._crear_tarjeta_kpi(kpi_row, "Nodos", f"{n}",
                                f"{len([1 for i in range(n) if nodos[i]['tipo']==1])} PQ · "
                                f"{len([1 for i in range(n) if nodos[i]['tipo']==2])} PV · "
                                f"{len([1 for i in range(n) if nodos[i]['tipo']==3])} Slack",
                                C['info'])

        nb = ttk.Notebook(ventana)
        nb.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0, 14))

        J_directo = datos['J']; m = J_directo.shape[0]
        incognitas = []
        for i in range(n):
            if nodos[i]['tipo'] != 3: incognitas.append(f"Δθ{i+1}")
        for i in idx_pq: incognitas.append(f"ΔV{i+1}")
        ecuaciones = []
        for i in range(n):
            if nodos[i]['tipo'] != 3: ecuaciones.append(f"ΔP{i+1}")
        for i in idx_pq: ecuaciones.append(f"ΔQ{i+1}")

        if num_iter == 0:
            mismatch_actual = datos['mismatch']
        else:
            mismatch_actual = datos.get('mismatch_actual')
            if mismatch_actual is None:
                tmp = []
                for i in range(n):
                    if nodos[i]['tipo'] != 3:
                        tmp.append(datos['deltaP'][i]
                                   if datos['deltaP'][i] is not None else 0.0)
                for i in idx_pq:
                    tmp.append(datos['deltaQ'][i]
                               if datos['deltaQ'][i] is not None else 0.0)
                mismatch_actual = np.array(tmp)
        delta_x = datos.get('delta_x', np.zeros(len(mismatch_actual)))

        # TAB 1: Estado de nodos
        tab1 = tk.Frame(nb, bg=C['panel'])
        nb.add(tab1, text='   Estado de Nodos   ')
        scroll1 = tk.Frame(tab1, bg=C['panel'])
        scroll1.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)
        self._crear_seccion(scroll1, "  ESTADO DE LOS NODOS", C['accent'])
        if num_iter == 0:
            headers = ["Nodo", "Tipo", "V (p.u.)", "θ (rad)",
                       "P_calc", "Q_calc", "ΔP", "ΔQ"]
        else:
            headers = ["Nodo", "Tipo", "V prev", "θ prev (rad)",
                       "V nuevo", "θ nuevo (rad)", "Δθ (rad)", "ΔV",
                       "P_calc", "Q_calc", "ΔP", "ΔQ"]
        rows = []
        for i in range(n):
            ts = {1:"PQ", 2:"PV", 3:"Slack"}.get(nodos[i]['tipo'], "PQ")
            dp = f"{deltaP[i]:.6e}" if deltaP[i] is not None else "—"
            dq = f"{deltaQ[i]:.6e}" if deltaQ[i] is not None else "—"
            if num_iter == 0:
                row = [i+1, ts, f"{V_actual[i]:.4f}", f"{theta_actual[i]:.5f}",
                       f"{P_calc[i]:+.6f}", f"{Q_calc[i]:+.6f}", dp, dq]
            else:
                row = [i+1, ts, f"{datos['V_prev'][i]:.4f}",
                       f"{datos['theta_prev'][i]:.5f}",
                       f"{V_actual[i]:.4f}", f"{theta_actual[i]:.5f}",
                       f"{datos['delta_theta_usado'][i]:+.5f}",
                       f"{datos['delta_V_usado'][i]:+.5f}",
                       f"{P_calc[i]:+.6f}", f"{Q_calc[i]:+.6f}", dp, dq]
            rows.append(row)

        def color_celda(r, c, v):
            tc = {1: C['pq'], 2: C['pv'], 3: C['slack']}
            if c == 1:
                for k, col in tc.items():
                    if {1:"PQ", 2:"PV", 3:"Slack"}[k] == v:
                        return (col, 'white')
            if num_iter == 0:
                if c in (6, 7) and v != "—" and abs(float(v)) > 0.01:
                    return (C['danger'], 'white')
            else:
                if c in (10, 11) and v != "—" and abs(float(v)) > 0.01:
                    return (C['danger'], 'white')
            return None

        self._crear_tabla_datos(scroll1, headers, rows,
                                 col_widths=[6, 8] + [10] * (len(headers) - 2),
                                 colorear_celdas=color_celda)

        # TAB 2: Jacobiano
        tab2 = tk.Frame(nb, bg=C['panel'])
        nb.add(tab2, text='   Jacobiano   ')
        canvas2 = tk.Canvas(tab2, bg=C['panel'], highlightthickness=0)
        sb2 = tk.Scrollbar(tab2, orient='vertical', command=canvas2.yview)
        inner2 = tk.Frame(canvas2, bg=C['panel'])
        inner2.bind('<Configure>',
                    lambda e: canvas2.configure(scrollregion=canvas2.bbox("all")))
        canvas2.create_window((0, 0), window=inner2, anchor='nw')
        canvas2.configure(yscrollcommand=sb2.set)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)
        canvas2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        c2 = tk.Frame(inner2, bg=C['panel'])
        c2.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        self._crear_seccion(c2, "  MATRIZ JACOBIANA (J)", C['pv'])
        info2 = tk.Frame(c2, bg=C['surface'], highlightthickness=1,
                          highlightbackground=C['border'])
        info2.pack(fill=tk.X, pady=(0, 16))
        tk.Label(info2,
                 text=f"Matriz de derivadas parciales (m × m) con m = {m}.\n"
                      f"Filas = ecuaciones de desbalance.  Columnas = incógnitas.",
                 font=font(9), bg=C['surface'], fg=C['text'],
                 justify='left', padx=14, pady=10).pack(anchor='w')
        if m > 0:
            mc = tk.Frame(c2, bg=C['panel']); mc.pack()
            hc = tk.Frame(mc, bg=C['panel']); hc.pack()
            tk.Label(hc, text="", width=8, bg=C['panel']).grid(row=0, column=0)
            tk.Label(hc, text="", width=2, bg=C['panel']).grid(row=0, column=1)
            for j, name in enumerate(incognitas):
                tk.Label(hc, text=name, font=font(9, "bold"),
                         bg=C['head_bg'], fg=C['accent'], width=10, pady=5
                         ).grid(row=0, column=j+2, padx=1, pady=1)
            br = tk.Frame(mc, bg=C['panel']); br.pack()
            nl = tk.Frame(br, bg=C['panel']); nl.pack(side=tk.LEFT, padx=(0, 4))
            tk.Label(nl, text="", bg=C['panel'], height=1).pack()
            for name in ecuaciones:
                tk.Label(nl, text=name, font=font(9, "bold"),
                         bg=C['head_bg'], fg=C['accent'], width=7, pady=6
                         ).pack(pady=1)
            J_str = [[f"{J_directo[i, j]:+.4f}" for j in range(m)] for i in range(m)]
            self._crear_matriz_bonita(br, J_str,
                                       highlight_diag=True).pack(side=tk.LEFT)

        # TAB 3: Sistema Lineal
        tab3 = tk.Frame(nb, bg=C['panel'])
        nb.add(tab3, text='   Sistema Lineal   ')
        canvas3 = tk.Canvas(tab3, bg=C['panel'], highlightthickness=0)
        sb3 = tk.Scrollbar(tab3, orient='vertical', command=canvas3.yview)
        inner3 = tk.Frame(canvas3, bg=C['panel'])
        inner3.bind('<Configure>',
                    lambda e: canvas3.configure(scrollregion=canvas3.bbox("all")))
        canvas3.create_window((0, 0), window=inner3, anchor='nw')
        canvas3.configure(yscrollcommand=sb3.set)
        sb3.pack(side=tk.RIGHT, fill=tk.Y)
        canvas3.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        c3 = tk.Frame(inner3, bg=C['panel'])
        c3.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        self._crear_seccion(c3,
                            "  RESOLUCIÓN DEL SISTEMA LINEAL  —  J · Δx = mismatch",
                            C['info'])
        desc = tk.Frame(c3, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        desc.pack(fill=tk.X, pady=(0, 18))
        tk.Label(desc, text="Se plantea el sistema lineal y se despeja el vector de incógnitas\n"
                            "multiplicando por la matriz inversa del Jacobiano.",
                 font=font(10), bg=C['surface'], fg=C['text'], justify='left',
                 padx=16, pady=12).pack(anchor='w')

        try:
            J_inv = np.linalg.inv(J_directo); J_inv_error = None
        except Exception as e:
            J_inv = None; J_inv_error = str(e)

        J_str = [[f"{J_directo[i, j]:+.4f}" for j in range(m)] for i in range(m)]
        dx_sym = incognitas[:]
        mis_str = [f"{mismatch_actual[i]:+.6f}" for i in range(m)]

        fila_sup = tk.Frame(c3, bg=C['panel'])
        fila_sup.pack(fill=tk.X, pady=(0, 20))
        col1 = tk.Frame(fila_sup, bg=C['panel'])
        col1.pack(side=tk.LEFT, padx=(0, 30))
        tk.Label(col1, text="①  Ecuación original", font=font(12, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(anchor='w', pady=(0, 8))
        sub1 = tk.Frame(col1, bg=C['panel']); sub1.pack()
        fJ = tk.Frame(sub1, bg=C['panel']); fJ.pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(fJ, text="J", font=font(10, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(pady=(0, 4))
        self._crear_matriz_bonita(fJ, J_str).pack()
        tk.Label(sub1, text="·", font=font(30, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=6)
        fdx = tk.Frame(sub1, bg=C['panel']); fdx.pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(fdx, text="Δx", font=font(10, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(pady=(0, 4))
        self._crear_vector_columna(fdx, dx_sym).pack()
        tk.Label(sub1, text="=", font=font(30, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=6)
        fmis = tk.Frame(sub1, bg=C['panel']); fmis.pack(side=tk.LEFT)
        tk.Label(fmis, text="mismatch", font=font(10, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(pady=(0, 4))
        self._crear_vector_columna(fmis, mis_str).pack()

        tk.Label(fila_sup, text="  ⟹  ", font=font(28, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=10)

        col2 = tk.Frame(fila_sup, bg=C['panel'])
        col2.pack(side=tk.LEFT, padx=(20, 0))
        tk.Label(col2, text="②  Despeje de las incógnitas", font=font(12, "bold"),
                 bg=C['panel'], fg=C['slack']).pack(anchor='w', pady=(0, 8))
        sub2 = tk.Frame(col2, bg=C['panel']); sub2.pack()
        fdx2 = tk.Frame(sub2, bg=C['panel']); fdx2.pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(fdx2, text="Δx", font=font(10, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(pady=(0, 4))
        self._crear_vector_columna(fdx2, dx_sym).pack()
        tk.Label(sub2, text="=", font=font(30, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=6)
        fJi = tk.Frame(sub2, bg=C['panel']); fJi.pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(fJi, text="J⁻¹", font=font(10, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(pady=(0, 4))
        if J_inv is not None:
            Ji_str = [[f"{J_inv[i, j]:+.2e}" for j in range(m)] for i in range(m)]
            self._crear_matriz_bonita(fJi, Ji_str).pack()
        else:
            tk.Label(fJi, text=f"[Error: {J_inv_error}]",
                     font=font(9, italic=True),
                     bg=C['panel'], fg=C['danger']).pack(padx=20, pady=20)
        tk.Label(sub2, text="·", font=font(30, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=6)
        fmis2 = tk.Frame(sub2, bg=C['panel']); fmis2.pack(side=tk.LEFT)
        tk.Label(fmis2, text="mismatch", font=font(10, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(pady=(0, 4))
        self._crear_vector_columna(fmis2, mis_str).pack()

        ley = tk.Frame(c3, bg=C['surface'], highlightthickness=1,
                       highlightbackground=C['border'])
        ley.pack(fill=tk.X, pady=(0, 16))
        tk.Label(ley, text="Filas del sistema (ecuaciones):  "
                           + "  ·  ".join(ecuaciones),
                 font=font_m(9), bg=C['surface'], fg=C['text']
                 ).pack(anchor='w', padx=14, pady=(10, 2))
        tk.Label(ley, text="Columnas / incógnitas del vector Δx:  "
                           + "  ·  ".join(incognitas),
                 font=font_m(9), bg=C['surface'], fg=C['text']
                 ).pack(anchor='w', padx=14, pady=(0, 10))

        res = tk.Frame(c3, bg=C['surface'], highlightthickness=1,
                       highlightbackground=C['border'])
        res.pack(fill=tk.X, pady=(0, 10))
        tk.Label(res,
                 text="③  Resultados de los incrementos (solución del sistema)",
                 font=font(12, "bold"), bg=C['surface'], fg=C['slack']
                 ).pack(anchor='w', padx=14, pady=(12, 8))
        for i, name in enumerate(incognitas):
            v = delta_x[i] if i < len(delta_x) else 0.0
            fila = tk.Frame(res, bg=C['surface'])
            fila.pack(anchor='w', padx=24, pady=1)
            tk.Label(fila, text=f"{name:>6s}  =", font=font_m(11, "bold"),
                     bg=C['surface'], fg=C['accent']).pack(side=tk.LEFT)
            unit = ' rad' if name.startswith('Δθ') else ' p.u.'
            tk.Label(fila, text=f"  {v:+.8f}{unit}", font=font_m(11, "bold"),
                     bg=C['surface'], fg=C['accent_l']).pack(side=tk.LEFT)
        tk.Label(res, text="", bg=C['surface'], height=1).pack()

        # TAB 4: Actualización
        tab4 = tk.Frame(nb, bg=C['panel'])
        nb.add(tab4, text='   Actualización   ')
        scroll4 = tk.Frame(tab4, bg=C['panel'])
        scroll4.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)
        self._crear_seccion(scroll4,
                            "  ACTUALIZACIÓN  (x⁽ⁱ⁺¹⁾ = x⁽ⁱ⁾ + Δx⁽ⁱ⁾)",
                            C['slack'])
        Vp_u = datos['V']; tp_u = datos['theta']
        dtu = datos['delta_theta']; dVu = datos['delta_V']
        tnu = tp_u + dtu; Vnu = Vp_u + dVu
        rows4 = []
        for i in range(n):
            if nodos[i]['tipo'] != 3:
                rows4.append([f"θ{i+1}", f"{tp_u[i]:+.6f} rad",
                              f"{dtu[i]:+.8f} rad",
                              f"{tnu[i]:+.6f} rad",
                              f"{tp_u[i]:.4f} + ({dtu[i]:+.6f}) = {tnu[i]:.4f}"])
            if nodos[i]['tipo'] == 1:
                rows4.append([f"V{i+1}", f"{Vp_u[i]:+.6f} p.u.",
                              f"{dVu[i]:+.8f} p.u.",
                              f"{Vnu[i]:+.6f} p.u.",
                              f"{Vp_u[i]:.4f} + ({dVu[i]:+.6f}) = {Vnu[i]:.4f}"])
        self._crear_tabla_datos(scroll4,
                                 ["Variable", "Valor anterior", "Corrección",
                                  "Valor nuevo", "Verificación"],
                                 rows4,
                                 col_widths=[10, 16, 16, 16, 32])

        # TAB 5: Ybus
        tab5 = tk.Frame(nb, bg=C['panel'])
        nb.add(tab5, text='   Matriz Ybus   ')
        canvas5 = tk.Canvas(tab5, bg=C['panel'], highlightthickness=0)
        sb5 = tk.Scrollbar(tab5, orient='vertical', command=canvas5.yview)
        inner5 = tk.Frame(canvas5, bg=C['panel'])
        inner5.bind('<Configure>',
                    lambda e: canvas5.configure(scrollregion=canvas5.bbox("all")))
        canvas5.create_window((0, 0), window=inner5, anchor='nw')
        canvas5.configure(yscrollcommand=sb5.set)
        sb5.pack(side=tk.RIGHT, fill=tk.Y)
        canvas5.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        c5 = tk.Frame(inner5, bg=C['panel'])
        c5.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        self._crear_seccion(c5, "  MATRIZ DE ADMITANCIA Ybus (p.u.)", C['accent'])
        info5 = tk.Frame(c5, bg=C['surface'], highlightthickness=1,
                          highlightbackground=C['border'])
        info5.pack(fill=tk.X, pady=(0, 16))
        tk.Label(info5,
                 text=f"Matriz compleja n × n con n = {n}. "
                      f"Elementos en formato a ± bj.",
                 font=font(9), bg=C['surface'], fg=C['text'],
                 justify='left', padx=14, pady=10).pack(anchor='w')

        mcy = tk.Frame(c5, bg=C['panel']); mcy.pack()
        hcy = tk.Frame(mcy, bg=C['panel']); hcy.pack()
        tk.Label(hcy, text="", width=8, bg=C['panel']).grid(row=0, column=0)
        tk.Label(hcy, text="", width=2, bg=C['panel']).grid(row=0, column=1)
        for j in range(n):
            tk.Label(hcy, text=f"N{j+1}", font=font(9, "bold"),
                     bg=C['head_bg'], fg=C['accent'], width=14, pady=5
                     ).grid(row=0, column=j+2, padx=1, pady=1)
        bry = tk.Frame(mcy, bg=C['panel']); bry.pack()
        nly = tk.Frame(bry, bg=C['panel']); nly.pack(side=tk.LEFT, padx=(0, 4))
        tk.Label(nly, text="", bg=C['panel'], height=1).pack()
        for i in range(n):
            tk.Label(nly, text=f"N{i+1}", font=font(9, "bold"),
                     bg=C['head_bg'], fg=C['accent'], width=7, pady=6
                     ).pack(pady=1)
        yb_str = []
        for i in range(n):
            fila = []
            for j in range(n):
                val = Ybus[i][j]
                if abs(val.imag) < 1e-12: fila.append(f"{val.real:+.4f}")
                elif abs(val.real) < 1e-12: fila.append(f"{val.imag:+.4f}j")
                else:
                    s = "+" if val.imag >= 0 else "-"
                    fila.append(f"{val.real:+.4f}{s}{abs(val.imag):.4f}j")
            yb_str.append(fila)
        self._crear_matriz_bonita(bry, yb_str).pack(side=tk.LEFT)

        pie = tk.Frame(ventana, bg=C['bg'])
        pie.pack(fill=tk.X, padx=14, pady=(0, 12))
        tk.Button(pie, text="Cerrar ventana", command=ventana.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=24, pady=8, cursor='hand2',
                  borderwidth=0,
                  activebackground=C['primary_l'],
                  activeforeground='white').pack(side=tk.RIGHT)

    def redisplay_iter(self, num_iter):
        if num_iter in self.datos_iteracion:
            self.mostrar_resultados_iterN(num_iter, self.datos_iteracion[num_iter])
        else:
            messagebox.showinfo("Info", f"Iteración {num_iter} no calculada.")

    # =================================================================
    # RESULTADOS FINALES
    # =================================================================
    def mostrar_resultados_finales(self):
        if not self.datos_iteracion:
            messagebox.showwarning("Sin resultados",
                "No hay iteraciones calculadas todavía.\n\n"
                "Presiona 'Calcular todo' primero.")
            return

        self._actualizar_bases()
        BASE_MVA = self.BASE_MVA
        V_BASE_KV = self.V_BASE_KV

        num_iter = max(self.datos_iteracion.keys())
        d = self.datos_iteracion[num_iter]
        n = d['n']
        nodos = d['nodos']
        V = d['V']
        theta = d['theta']
        P_calc = d['P_calc']
        Q_calc = d['Q_calc']

        filas = []
        P_gen_total = Q_gen_total = 0.0
        P_load_total = Q_load_total = 0.0

        for i in range(n):
            tipo = nodos[i]['tipo']
            V_i = V[i]
            th_deg = math.degrees(theta[i])
            Pc_pu = P_calc[i]
            Qc_pu = Q_calc[i]

            Pd_real = self._parse_float(self.Pd_widgets[i].get(), 0.0)
            Qd_real = self._parse_float(self.Qd_widgets[i].get(), 0.0)

            P_gen = Pd_real + Pc_pu * BASE_MVA
            Q_gen = Qd_real + Qc_pu * BASE_MVA

            P_load = Pd_real
            Q_load = Qd_real

            P_gen_total += P_gen
            Q_gen_total += Q_gen
            P_load_total += P_load
            Q_load_total += Q_load

            filas.append({
                'nodo': i + 1,
                'tipo': {1: "PQ", 2: "PV", 3: "Slack"}[tipo],
                'V_kv': V_i * V_BASE_KV,
                'theta_deg': th_deg,
                'P_calc': Pc_pu * BASE_MVA,
                'Q_calc': Qc_pu * BASE_MVA,
                'P_gen': P_gen,
                'Q_gen': Q_gen,
                'P_load': P_load,
                'Q_load': Q_load,
            })

        P_loss = P_gen_total - P_load_total
        Q_loss = Q_gen_total - Q_load_total

        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Resultados Finales — Vatia")
        win.geometry("1180x820")
        win.minsize(1000, 640)
        win.configure(bg=C['bg'])
        win.transient(self.root)

        hdr = tk.Frame(win, bg=C['primary_d'], height=76)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['accent'], height=4).pack(fill=tk.X, side=tk.TOP)

        tk.Label(hdr, text="  RESULTADOS FINALES DEL FLUJO DE POTENCIA",
                 font=font(15, "bold"), bg=C['primary_d'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=22)

        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['danger'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=14, pady=6, cursor='hand2',
                  borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=22)

        badge = tk.Frame(hdr, bg=C['accent'])
        badge.pack(side=tk.RIGHT, padx=6, pady=22)
        tk.Label(badge, text=f"  Última iteración: {num_iter}  ",
                 font=font(10, "bold"), bg=C['accent'], fg='white',
                 padx=10, pady=6).pack()

        canvas = tk.Canvas(win, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(win, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas.find_all()[0],
                                                 width=e.width))
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        def _wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        canvas.bind('<MouseWheel>', _wheel)
        inner.bind('<MouseWheel>', _wheel)

        tk.Label(inner, text="BALANCE DE POTENCIA DEL SISTEMA",
                 font=font(13, "bold"), bg=C['bg'], fg=C['primary']
                 ).pack(anchor='w', padx=30, pady=(22, 12))

        info_base = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                             highlightbackground=C['border'])
        info_base.pack(fill=tk.X, padx=30, pady=(0, 14))
        tk.Label(info_base,
                 text=f"Bases del sistema:    S_base = {BASE_MVA:.1f} MVA    ·    "
                      f"V_base = {V_BASE_KV:.1f} kV    ·    "
                      f"Z_base = {V_BASE_KV**2 / BASE_MVA:.2f} Ω",
                 font=font_m(9), bg=C['surface'], fg=C['text'],
                 padx=14, pady=8, anchor='w').pack(fill=tk.X)

        kpis = tk.Frame(inner, bg=C['bg'])
        kpis.pack(fill=tk.X, padx=30, pady=(0, 22))

        def _kpi(parent, titulo, valor_p, valor_q, color):
            card = tk.Frame(parent, bg=C['panel'], highlightthickness=1,
                            highlightbackground=C['border'])
            card.pack(side=tk.LEFT, padx=6, fill=tk.Y)
            tk.Frame(card, bg=color, width=5).pack(side=tk.LEFT, fill=tk.Y)
            body = tk.Frame(card, bg=C['panel'])
            body.pack(side=tk.LEFT, padx=16, pady=14)
            tk.Label(body, text=titulo.upper(),
                     font=font(8, "bold"), bg=C['panel'], fg=C['text_l']
                     ).pack(anchor='w')
            tk.Label(body, text=f"P = {valor_p:+.2f} MW",
                     font=font(13, "bold"), bg=C['panel'], fg=color
                     ).pack(anchor='w', pady=(4, 0))
            tk.Label(body, text=f"Q = {valor_q:+.2f} MVAr",
                     font=font(11), bg=C['panel'], fg=C['text_l']
                     ).pack(anchor='w')

        _kpi(kpis, "Generación total",
             P_gen_total, Q_gen_total, C['slack'])
        _kpi(kpis, "Demanda total",
             P_load_total, Q_load_total, C['pq'])
        _kpi(kpis, "Pérdidas del sistema",
             P_loss, Q_loss, C['danger'])

        tk.Label(inner, text="ESTADO FINAL DE LOS NODOS",
                 font=font(13, "bold"), bg=C['bg'], fg=C['primary']
                 ).pack(anchor='w', padx=30, pady=(0, 12))

        tabla_frame = tk.Frame(inner, bg=C['border'])
        tabla_frame.pack(fill=tk.X, padx=30, pady=(0, 20))

        headers = ["Nodo", "Tipo", "V (kV)", "θ (°)",
                   "P_calc (MW)", "Q_calc (MVAr)",
                   "Pg (MW)", "Qg (MVAr)",
                   "Pd (MW)", "Qd (MVAr)"]

        head = tk.Frame(tabla_frame, bg=C['primary_d'])
        head.pack(fill=tk.X)
        for i, h in enumerate(headers):
            head.grid_columnconfigure(i, weight=1, uniform='col')
            tk.Label(head, text=h, font=font(9, "bold"),
                     bg=C['primary_d'], fg='white',
                     padx=8, pady=8
                     ).grid(row=0, column=i, sticky='nsew', padx=1, pady=1)

        body = tk.Frame(tabla_frame, bg=C['border'])
        body.pack(fill=tk.X)
        for r_idx, fila in enumerate(filas):
            bg_row = C['row_even'] if r_idx % 2 == 0 else C['row_odd']
            celdas = [
                f"{fila['nodo']}",
                fila['tipo'],
                f"{fila['V_kv']:.2f}",
                f"{fila['theta_deg']:+.4f}",
                f"{fila['P_calc']:+.2f}",
                f"{fila['Q_calc']:+.2f}",
                f"{fila['P_gen']:+.2f}",
                f"{fila['Q_gen']:+.2f}",
                f"{fila['P_load']:+.2f}",
                f"{fila['Q_load']:+.2f}",
            ]
            for c_idx, val in enumerate(celdas):
                body.grid_columnconfigure(c_idx, weight=1, uniform='col')
                cell_bg = bg_row
                cell_fg = C['text']
                if c_idx == 1:
                    colors_tipo = {'Slack': C['slack'], 'PV': C['pv'],
                                   'PQ': C['pq']}
                    cell_bg = colors_tipo.get(fila['tipo'], bg_row)
                    cell_fg = 'white'
                tk.Label(body, text=val, font=font_m(9),
                         bg=cell_bg, fg=cell_fg,
                         padx=8, pady=6
                         ).grid(row=r_idx, column=c_idx, sticky='nsew',
                                padx=1, pady=1)

        note = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        note.pack(fill=tk.X, padx=30, pady=(0, 20))
        tk.Label(note, text="NOTA SOBRE EL CÁLCULO",
                 font=font(9, "bold"), bg=C['surface'], fg=C['text'],
                 anchor='w').pack(fill=tk.X, padx=14, pady=(10, 2))
        tk.Label(note,
                 text="• Fórmula general:  P_gen[i] = Pd[i] + P_calc[i]·S_base\n"
                      "     Es decir, la generación de cada nodo = lo que demanda + lo que entrega al sistema.\n"
                      "• Los valores Pg / Qg de nodos Slack y PV son calculados (no los ingresados).\n"
                      "• Los valores Pd / Qd y Pg / Qg de nodos PQ son los que ingresaste.\n"
                      "• Pérdidas del sistema = Generación total − Demanda total (en MW y MVAr).",
                 font=font_m(9), bg=C['surface'], fg=C['text_l'],
                 justify='left', anchor='w').pack(fill=tk.X, padx=14, pady=(0, 12))

        tk.Button(inner, text="Cerrar ventana", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=30, pady=10, cursor='hand2',
                  borderwidth=0).pack(pady=20)

    # =================================================================
    # DIAGRAMA UNIFILAR GRANDE
    # =================================================================
    def mostrar_diagrama_unifilar(self):
        try:
            n = int(self.entry_n.get())
        except Exception:
            messagebox.showerror("Error", "Número de nodos inválido.")
            return
        def _f(w, default=0.0):
            return self._parse_float(w.get(), default)

        nodos = []
        for i in range(n):
            ts = self.tipo_widgets[i].get()
            tipo = 3 if "Slack" in ts else (2 if "PV" in ts else 1)
            Pd_i = _f(self.Pd_widgets[i], 0.0)
            Qd_i = _f(self.Qd_widgets[i], 0.0)
            Pg_i = _f(self.Pg_widgets[i], 0.0)
            Qg_i = _f(self.Qg_widgets[i], 0.0)
            nodos.append({'tipo': tipo,
                          'V': _f(self.V_fijo_widgets[i], 1.0),
                          'theta': _f(self.theta_fijo_widgets[i], 0.0),
                          'Pesp': (Pg_i - Pd_i) / self.BASE_MVA,
                          'Qesp': (Qg_i - Qd_i) / self.BASE_MVA,
                          'Pd_MW': Pd_i, 'Qd_MVAr': Qd_i,
                          'Pg_MW': Pg_i, 'Qg_MVAr': Qg_i})
        lineas = []
        for row in self.entries_lineas:
            a = self._parse_int(row[0].get(), 0); b = self._parse_int(row[1].get(), 0)
            G = self._parse_float(row[2].get(), 0.0); B = self._parse_float(row[3].get(), 0.0)
            ysh = self._parse_float(row[4].get(), 0.0)
            if a > 0 and b > 0 and a <= n and b <= n and a != b:
                lineas.append({'a': a-1, 'b': b-1, 'G': G, 'B': B, 'y_shunt': ysh})

        BG_BG       = "#EDF1F6"; BG_TOOLBAR = "#1E3A5F"; BG_CANVAS = "#FFFFFF"
        GRID_FINE   = "#F1F5F9"; GRID_BOLD = "#E2E8F0"
        TITLE_COL   = "#1E3A5F"; SUBTITLE = "#64748B"
        LINE_DARK   = "#475569"; LINE_SHADOW = "#CBD5E1"
        LABEL_BG    = "#FFFFFF"; LABEL_BORDER = "#0E7490"; LABEL_TEXT = "#0E7490"
        ARROW_COLOR = "#C2410C"
        SLACK_COL   = "#047857"; PV_COL = "#6D28D9"; PQ_COL = "#C2410C"
        SLACK_RING  = "#D1FAE5"; PV_RING = "#EDE9FE"; PQ_RING = "#FFEDD5"
        INFO_BG     = "#FFFFFF"

        ventana = tk.Toplevel(self.root)
        ventana.title("Diagrama Unifilar — Vatia")
        ventana.geometry("1250x850")
        ventana.configure(bg=BG_BG)

        toolbar = tk.Frame(ventana, bg=BG_TOOLBAR, height=58)
        toolbar.pack(fill=tk.X, side=tk.TOP); toolbar.pack_propagate(False)
        tk.Frame(toolbar, bg="#0E7490", height=4).pack(fill=tk.X, side=tk.TOP)
        tk.Label(toolbar, text="DIAGRAMA UNIFILAR",
                 font=font(15, "bold"), bg=BG_TOOLBAR, fg="white"
                 ).pack(side=tk.LEFT, padx=20, pady=14)
        tk.Label(toolbar, text="Arrastra para mover  ·  Ctrl + rueda para zoom",
                 font=font(9, italic=True), bg=BG_TOOLBAR, fg="#CBD5E1"
                 ).pack(side=tk.LEFT, pady=14)

        main = tk.Frame(ventana, bg=BG_BG)
        main.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        canvas = tk.Canvas(main, bg=BG_CANVAS, highlightthickness=1,
                            highlightbackground="#CBD5E1")
        hbar = tk.Scrollbar(main, orient=tk.HORIZONTAL, command=canvas.xview)
        vbar = tk.Scrollbar(main, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def on_press(e): canvas.scan_mark(e.x, e.y)
        def on_drag(e): canvas.scan_dragto(e.x, e.y, gain=1)
        def on_wheel(e):
            if not (e.state & 0x0004):
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"); return
            x = canvas.canvasx(e.x); y = canvas.canvasy(e.y)
            factor = 1.1 if e.delta > 0 else 0.9
            canvas.scale("all", x, y, factor, factor)
            canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<MouseWheel>", on_wheel)

        def reset_view():
            canvas.xview_moveto(0); canvas.yview_moveto(0)
        def ajustar_vista():
            canvas.update_idletasks()
            bb = canvas.bbox("all")
            if bb is None: return
            canvas.configure(scrollregion=bb)
            total_w = bb[2] - bb[0]; total_h = bb[3] - bb[1]
            vis_w = canvas.winfo_width(); vis_h = canvas.winfo_height()
            if total_w > vis_w:
                canvas.xview_moveto((total_w - vis_w) / 2 / total_w)
            if total_h > vis_h:
                canvas.yview_moveto((total_h - vis_h) / 2 / total_h)

        tk.Button(toolbar, text="⟲  Reset Vista", command=reset_view,
                  bg="#334155", fg="white", relief=tk.FLAT,
                  font=font(9), cursor="hand2", padx=14, pady=4
                  ).pack(side=tk.RIGHT, padx=8, pady=12)
        tk.Button(toolbar, text="⊡  Ajustar Vista", command=ajustar_vista,
                  bg="#047857", fg="white", relief=tk.FLAT,
                  font=font(9), cursor="hand2", padx=14, pady=4
                  ).pack(side=tk.RIGHT, padx=4, pady=12)

        R_layout = 340; margen = 260
        CW = 2 * R_layout + 2 * margen; CH = 2 * R_layout + 2 * margen
        cx, cy = CW / 2, CH / 2
        posiciones = []
        for i in range(n):
            ang = 2 * math.pi * i / n - math.pi / 2
            posiciones.append((cx + R_layout * math.cos(ang),
                                cy + R_layout * math.sin(ang)))

        for x in range(0, CW, 40):
            canvas.create_line(x, 0, x, CH, fill=GRID_FINE, width=1)
        for y in range(0, CH, 40):
            canvas.create_line(0, y, CW, y, fill=GRID_FINE, width=1)
        for x in range(0, CW, 200):
            canvas.create_line(x, 0, x, CH, fill=GRID_BOLD, width=1)
        for y in range(0, CH, 200):
            canvas.create_line(0, y, CW, y, fill=GRID_BOLD, width=1)

        canvas.create_text(CW/2, 40, text="DIAGRAMA UNIFILAR DEL SISTEMA",
                           font=font(22, "bold"), fill=TITLE_COL)
        canvas.create_text(CW/2, 72,
                           text=f"Sistema de {n} nodos  •  {len(lineas)} línea(s) de transmisión",
                           font=font(11, italic=True), fill=SUBTITLE)

        V_complex_calc = None
        if self.datos_iteracion and 0 in self.datos_iteracion:
            try:
                d0 = self.datos_iteracion[0]
                if d0.get('n') == n:
                    V_complex_calc = [d0['V'][i] * complex(math.cos(d0['theta'][i]),
                                                            math.sin(d0['theta'][i]))
                                      for i in range(n)]
            except Exception:
                V_complex_calc = None

        for line in lineas:
            i, j = line['a'], line['b']
            x1, y1 = posiciones[i]; x2, y2 = posiciones[j]
            dx, dy = x2 - x1, y2 - y1
            d = math.hypot(dx, dy) or 1.0
            ux, uy = dx / d, dy / d
            offset = 55
            sx1, sy1 = x1 + ux * offset, y1 + uy * offset
            sx2, sy2 = x2 - ux * offset, y2 - uy * offset

            canvas.create_line(sx1+3, sy1+3, sx2+3, sy2+3, width=5,
                               fill=LINE_SHADOW, capstyle=tk.ROUND)
            canvas.create_line(sx1, sy1, sx2, sy2, width=4,
                               fill=LINE_DARK, capstyle=tk.ROUND)
            for (px, py, px2, py2) in [(x1,y1,sx1,sy1), (sx2,sy2,x2,y2)]:
                canvas.create_line(px, py, px2, py2, width=2.5, fill=LINE_DARK)

            P_ij = None
            if V_complex_calc is not None:
                try:
                    Va = V_complex_calc[i]; Vb = V_complex_calc[j]
                    I_ij = ((Va - Vb) * complex(line['G'], line['B'])
                             + Va * complex(0, line.get('y_shunt', 0)))
                    P_ij = (Va * I_ij.conjugate()).real
                except Exception:
                    P_ij = None

            if P_ij is not None:
                fwd = P_ij >= 0
                t_pos = 0.22
                ax0 = sx1 + (sx2-sx1)*t_pos; ay0 = sy1 + (sy2-sy1)*t_pos
                adx = (sx2-sx1) if fwd else (sx1-sx2)
                ady = (sy2-sy1) if fwd else (sy1-sy2)
                alen = math.hypot(adx, ady) or 1.0
                aux, auy = adx/alen, ady/alen
                px, py = -auy, aux
                sz, hw = 20, 11
                canvas.create_polygon(ax0+aux*sz/2, ay0+auy*sz/2,
                                       ax0-aux*sz/2+px*hw, ay0-auy*sz/2+py*hw,
                                       ax0-aux*sz/2-px*hw, ay0-auy*sz/2-py*hw,
                                       fill=ARROW_COLOR, outline="white",
                                       width=1.5)

            G, B = line['G'], line['B']
            if P_ij is not None:
                label = f"Y = {G:+.4f} {'+' if B>=0 else '-'} j{abs(B):.4f}\nP = {P_ij:+.4f} p.u."
                fuente_lbl = font_m(9, "bold"); pad_y = 6
            else:
                label = f"Y = {G:+.4f} {'+' if B>=0 else '-'} j{abs(B):.4f}"
                fuente_lbl = font_m(10, "bold"); pad_y = 4

            mx, my = (sx1+sx2)/2, (sy1+sy2)/2
            tid = canvas.create_text(mx, my, text=label, font=fuente_lbl,
                                     fill=LABEL_TEXT, justify=tk.CENTER)
            bb = canvas.bbox(tid)
            if bb:
                canvas.create_rectangle(bb[0]-8, bb[1]-pad_y,
                                         bb[2]+8, bb[3]+pad_y,
                                         fill=LABEL_BG, outline=LABEL_BORDER,
                                         width=1.5)
                canvas.tag_raise(tid)

        BUS_LEN = 90
        for i, (x, y) in enumerate(posiciones):
            tipo = nodos[i]['tipo']
            color = {3: SLACK_COL, 2: PV_COL, 1: PQ_COL}[tipo]
            ring = {3: SLACK_RING, 2: PV_RING, 1: PQ_RING}[tipo]

            canvas.create_line(x-BUS_LEN/2, y, x+BUS_LEN/2, y, width=13,
                               fill="#1F2937", capstyle=tk.ROUND)
            canvas.create_oval(x-22, y-22, x+22, y+22, fill=ring, outline='')
            canvas.create_oval(x-18, y-18, x+18, y+18, fill=color,
                               outline="white", width=3)
            canvas.create_text(x, y, text=str(i+1),
                               font=font(13, "bold"), fill="white")

            if tipo in (2, 3):
                gr = 26; gy = y - 78
                canvas.create_line(x, y-22, x, gy+gr, width=2.5, fill=LINE_DARK)
                canvas.create_oval(x-gr, gy-gr, x+gr, gy+gr, fill="#FFFFFF",
                                   outline=SLACK_COL, width=3)
                pts = []
                for t in range(31):
                    px = x-17+34*t/30; py = gy + 8*math.sin(2*math.pi*t/30)
                    pts.extend([px, py])
                canvas.create_line(pts, smooth=True, width=2.5, fill=SLACK_COL)
                canvas.create_text(x+gr+14, gy, text=f"G{i+1}",
                                   font=font(10, "bold"), fill=SLACK_COL,
                                   anchor="w")

            if tipo == 1:
                ly = y + 78
                canvas.create_line(x, y+22, x, ly-18, width=2.5, fill=LINE_DARK)
                canvas.create_polygon(x-12, ly-18, x+12, ly-18, x, ly,
                                       fill=PQ_COL, outline="#9A3412", width=1.5)
                canvas.create_text(x+24, ly-6, text=f"Carga {i+1}",
                                   font=font(10, "bold"), fill=PQ_COL,
                                   anchor="w")

            dxo = x - cx; dyo = y - cy
            dist = math.hypot(dxo, dyo) or 1.0
            uxo, uyo = dxo/dist, dyo/dist
            tx = x + uxo*145; ty = y + uyo*125
            tipo_txt = {3:"Slack", 2:"PV", 1:"PQ"}[tipo]
            lineas_info = [f"Nodo {i+1}  [{tipo_txt}]",
                           f"V  = {nodos[i]['V']:.3f} p.u.",
                           f"θ  = {nodos[i]['theta']:.2f} °"]
            if tipo == 1:
                lineas_info.append(f"Pd = {nodos[i]['Pd_MW']:.2f} MW")
                lineas_info.append(f"Qd = {nodos[i]['Qd_MVAr']:.2f} MVAr")
                if abs(nodos[i]['Pg_MW']) > 1e-9:
                    lineas_info.append(f"Pg = {nodos[i]['Pg_MW']:.2f} MW")
                if abs(nodos[i]['Qg_MVAr']) > 1e-9:
                    lineas_info.append(f"Qg = {nodos[i]['Qg_MVAr']:.2f} MVAr")
            elif tipo == 2:
                lineas_info.append(f"Pg = {nodos[i]['Pg_MW']:.2f} MW")
                lineas_info.append("Q  :  libre")
            else:
                lineas_info.append("V  :  fija (ref.)")
                lineas_info.append("θ  :  fija (0°)")
            info = "\n".join(lineas_info)

            tid = canvas.create_text(tx, ty, text=info, font=font_m(9),
                                     fill="#1F2937", justify=tk.LEFT)
            bb = canvas.bbox(tid)
            if bb:
                canvas.create_rectangle(bb[0]-10, bb[1]-7, bb[2]+10, bb[3]+7,
                                         fill=INFO_BG, outline=color, width=2)
                canvas.tag_raise(tid)

        lx0, ly0 = 25, CH - 240
        lx1, ly1 = lx0 + 300, ly0 + 220
        canvas.create_rectangle(lx0+4, ly0+4, lx1+4, ly1+4,
                                fill="#CBD5E1", outline="")
        canvas.create_rectangle(lx0, ly0, lx1, ly1, fill="#FFFFFF",
                                outline="#1E3A5F", width=2)
        canvas.create_text((lx0+lx1)/2, ly0+22, text="LEYENDA",
                           font=font(12, "bold"), fill="#1E3A5F")
        canvas.create_line(lx0+40, ly0+38, lx1-40, ly0+38,
                           fill="#CBD5E1", width=1)
        items = [(SLACK_COL, "Slack (barra infinita)"),
                 (PV_COL, "PV  (generador)"),
                 (PQ_COL, "PQ  (carga)")]
        for k, (col, txt) in enumerate(items):
            yk = ly0 + 58 + k*28
            canvas.create_oval(lx0+20, yk-9, lx0+38, yk+9, fill=col,
                               outline="white", width=2)
            canvas.create_text(lx0+50, yk, text=txt, font=font(10),
                               fill="#1F2937", anchor="w")
        yk = ly0 + 58 + 3*28
        canvas.create_line(lx0+20, yk, lx0+70, yk, width=8,
                           fill="#1F2937", capstyle=tk.ROUND)
        canvas.create_text(lx0+82, yk, text="Barra (nodo)",
                           font=font(10), fill="#1F2937", anchor="w")
        yk += 28
        canvas.create_line(lx0+20, yk, lx0+70, yk, width=4, fill=LINE_DARK)
        canvas.create_text(lx0+82, yk, text="Línea de transmisión",
                           font=font(10), fill="#1F2937", anchor="w")

        ventana.after(80, ajustar_vista)
        tk.Button(ventana, text="Cerrar", command=ventana.destroy,
                  bg="#B91C1C", fg="white", relief=tk.FLAT,
                  font=font(10, "bold"), width=15, cursor="hand2"
                  ).pack(pady=8)

    # =================================================================
    # GRÁFICA DE CONVERGENCIA
    # =================================================================
    def _show_convergence_chart(self):
        if not self.datos_iteracion:
            messagebox.showinfo("Info", "Calcula primero alguna iteración.")
            return
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Convergencia — Vatia")
        win.geometry("900x600"); win.configure(bg=C['bg'])
        hdr = tk.Frame(win, bg=C['primary'], height=64)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="  ANÁLISIS DE CONVERGENCIA",
                 font=font(15, "bold"), bg=C['primary'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=16)
        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['danger'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=14, pady=6,
                  borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=14)
        cvs = tk.Canvas(win, bg=C['panel'], highlightthickness=0)
        cvs.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        iters, err_p, err_q = [], [], []
        for k in sorted(self.datos_iteracion.keys()):
            d = self.datos_iteracion[k]
            dp = [abs(x) for x in d['deltaP'] if x is not None]
            dq = [abs(x) for x in d['deltaQ'] if x is not None]
            iters.append(k)
            err_p.append(max(dp) if dp else 0.0)
            err_q.append(max(dq) if dq else 0.0)
        if not iters:
            cvs.create_text(450, 250, text="Sin datos", fill=C['text_l'],
                            font=font(14, "bold"))
            return
        W = 820; H = 460
        cvs.configure(scrollregion=(0, 0, W + 40, H + 40))
        pad_l, pad_r, pad_t, pad_b = 80, 40, 60, 60
        plot_w = W - pad_l - pad_r; plot_h = H - pad_t - pad_b
        cvs.create_text(W/2 + 20, 25,
                        text="Error máximo por iteración (escala log)",
                        font=font(13, "bold"), fill=C['text'])
        x0, y0 = pad_l, pad_t + plot_h
        cvs.create_line(x0, pad_t, x0, y0, fill=C['text_l'], width=1)
        cvs.create_line(x0, y0, pad_l + plot_w, y0, fill=C['text_l'], width=1)
        all_errs = [e for e in err_p + err_q if e > 0] or [1e-6]
        log_min = min(math.floor(math.log10(min(all_errs))), -3)
        log_max = max(math.ceil(math.log10(max(all_errs))) + 1, 0)
        n_lines = log_max - log_min + 1
        for k in range(n_lines):
            v = 10 ** (log_max - k)
            y = pad_t + (k / (n_lines - 1)) * plot_h if n_lines > 1 else pad_t
            cvs.create_line(x0 - 5, y, pad_l + plot_w, y,
                            fill=C['border'], width=1, dash=(2, 4))
            cvs.create_text(x0 - 10, y, text=f"{v:.0e}", font=font_m(9),
                            fill=C['text_l'], anchor='e')
        n_pts = len(iters)
        step_x = plot_w / max(1, n_pts - 1) if n_pts > 1 else 0
        for i, it in enumerate(iters):
            x = x0 + i * step_x
            cvs.create_text(x, y0 + 18, text=f"Iter {it}",
                            font=font(10, "bold"), fill=C['text'])
        def to_y(v):
            if v <= 0: v = 1e-15
            lv = math.log10(v)
            t = (log_max - lv) / (log_max - log_min)
            return pad_t + t * plot_h
        for name, errs, color in [("ΔP", err_p, C['accent']),
                                    ("ΔQ", err_q, C['pv'])]:
            pts = []
            for i, e in enumerate(errs):
                x = x0 + i * step_x; y = to_y(e); pts.extend([x, y])
            if len(pts) >= 4:
                cvs.create_line(*pts, fill=color, width=2.5,
                                capstyle=tk.ROUND, smooth=True)
            for i, e in enumerate(errs):
                x = x0 + i * step_x; y = to_y(e)
                cvs.create_oval(x - 5, y - 5, x + 5, y + 5,
                                fill=color, outline='white', width=2)
        lx = pad_l + plot_w - 140; ly = pad_t + 10
        cvs.create_rectangle(lx, ly, lx + 130, ly + 60,
                             fill=C['surface'], outline=C['border'], width=1)
        cvs.create_line(lx + 10, ly + 20, lx + 40, ly + 20,
                        fill=C['accent'], width=3)
        cvs.create_text(lx + 50, ly + 20, text="máx |ΔP|",
                        font=font(9), fill=C['text'], anchor='w')
        cvs.create_line(lx + 10, ly + 42, lx + 40, ly + 42,
                        fill=C['pv'], width=3)
        cvs.create_text(lx + 50, ly + 42, text="máx |ΔQ|",
                        font=font(9), fill=C['text'], anchor='w')
        if log_min <= -3 <= log_max:
            yt = to_y(1e-3)
            cvs.create_line(x0, yt, pad_l + plot_w, yt,
                            fill=C['success'], width=2, dash=(6, 3))
            cvs.create_text(pad_l + plot_w - 5, yt - 10,
                            text="tol = 1e-3", font=font(9, "bold"),
                            fill=C['success'], anchor='e')
        tk.Button(win, text="Cerrar", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=24, pady=8, cursor='hand2',
                  borderwidth=0).pack(pady=(0, 14))

    # =================================================================
    # AYUDA
    # =================================================================
    def _open_help(self):
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Ayuda — Vatia")
        win.geometry("900x700"); win.configure(bg=C['bg'])
        hdr = tk.Frame(win, bg=C['primary'], height=64)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="  AYUDA Y FORMULARIO",
                 font=font(15, "bold"), bg=C['primary'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=16)
        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['danger'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=14, pady=6,
                  borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=14)
        cvs = tk.Canvas(win, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(win, orient='vertical', command=cvs.yview)
        inner = tk.Frame(cvs, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: cvs.configure(scrollregion=cvs.bbox("all")))
        cvs.create_window((0, 0), window=inner, anchor='nw')
        cvs.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        cvs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def seccion(titulo, color):
            f = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                          highlightbackground=C['border'])
            f.pack(fill=tk.X, padx=20, pady=(20, 8))
            tk.Frame(f, bg=color, width=5).pack(side=tk.LEFT, fill=tk.Y)
            tk.Label(f, text=titulo, font=font(12, "bold"),
                     bg=C['surface'], fg=C['text'], padx=14, pady=8
                     ).pack(side=tk.LEFT)
        def texto(txt):
            tk.Label(inner, text=txt, font=font_m(10), bg=C['bg'],
                     fg=C['text'], justify='left', anchor='w'
                     ).pack(fill=tk.X, padx=30, pady=2)

        seccion("1. ECUACIONES DE POTENCIA NODAL", C['accent'])
        texto("S_i = V_i · I_i*    donde    I_i = Σ Y_ij · V_j")
        texto("P_i = V_i · Σ V_j (G_ij cos θ_ij + B_ij sin θ_ij)")
        texto("Q_i = V_i · Σ V_j (G_ij sin θ_ij - B_ij cos θ_ij)")
        seccion("2. CLASIFICACIÓN DE NODOS", C['slack'])
        texto("• Slack:  V y θ conocidos  →  P y Q se calculan")
        texto("• PV:     P y |V| conocidos  →  Q y θ se calculan")
        texto("• PQ:     P y Q conocidos  →  |V| y θ se calculan")
        seccion("3. MISMATCH (DESBALANCE)", C['pv'])
        texto("ΔP_i = P_esp,i - P_calc,i     (para nodos PQ y PV)")
        texto("ΔQ_i = Q_esp,i - Q_calc,i     (solo para nodos PQ)")
        seccion("4. MATRIZ JACOBIANA", C['pq'])
        texto("[ ΔP ]   [ H  N ]   [ Δθ ]")
        texto("[ ΔQ ] = [ M  L ] · [ ΔV ]")
        texto("H_ii = -Q_i - B_ii·V_i²       H_ij = V_i·V_j·(G_ij sinθij - B_ij cosθij)")
        texto("N_ii = P_i/V_i + G_ii·V_i      N_ij = V_i·(G_ij cosθij + B_ij sinθij)")
        texto("M_ii = P_i - G_ii·V_i²         M_ij = -V_i·V_j·(G_ij cosθij + B_ij sinθij)")
        texto("L_ii = Q_i/V_i - B_ii·V_i      L_ij = V_i·(G_ij sinθij - B_ij cosθij)")
        seccion("5. PROCESO ITERATIVO", C['info'])
        texto("① Evaluar P_calc, Q_calc con V, θ actuales")
        texto("② Calcular mismatch (ΔP, ΔQ)")
        texto("③ Verificar convergencia: max|ΔP|, max|ΔQ| < ε")
        texto("④ Ensamblar Jacobiano J(V, θ)")
        texto("⑤ Resolver: J · Δx = mismatch   →   Δx = J⁻¹ · mismatch")
        texto("⑥ Actualizar: θ⁽ⁱ⁺¹⁾ = θ⁽ⁱ⁾ + Δθ ;  V⁽ⁱ⁺¹⁾ = V⁽ⁱ⁾ + ΔV")
        texto("⑦ Volver al paso ①")
        seccion("6. ATAJOS DE TECLADO", C['warning'])
        texto("Tab            →  Siguiente campo")
        texto("Shift+Tab      →  Campo anterior")
        texto("Enter          →  Confirmar y avanzar")
        texto("Ctrl + rueda   →  Zoom en diagramas")
        texto("F1             →  Abrir esta ayuda")
        tk.Button(inner, text="Cerrar", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=24, pady=8, cursor='hand2',
                  borderwidth=0).pack(pady=20)

    # =================================================================
    # ACERCA DE
    # =================================================================
    def _acerca_de(self):
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Acerca de Vatia")
        win.geometry("460x430"); win.configure(bg=C['bg'])
        win.transient(self.root); win.grab_set()
        frame = tk.Frame(win, bg=C['panel'])
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        cv = tk.Canvas(frame, width=90, height=90, bg=C['panel'],
                       highlightthickness=0)
        cv.pack(pady=(10, 10))
        draw_vatia_logo(cv, 45, 45, size=82, style='nexus',
                         bg_color=C['primary'], fg_color='#FFFFFF',
                         accent_color=C['accent'], slack_color=C['slack'],
                         pv_color=C['pv'], pq_color=C['pq'])
        tk.Label(frame, text="V A T I A", font=font(22, "bold"),
                 bg=C['panel'], fg=C['primary']).pack()
        tk.Label(frame, text="Suite de Análisis de Sistemas de Potencia",
                 font=font(10), bg=C['panel'], fg=C['text_l']).pack()
        tk.Label(frame, text="Versión 1.0.0", font=font_m(9),
                 bg=C['panel'], fg=C['text_ll']).pack(pady=(10, 4))
        tk.Label(frame, text="Desarrollado por Kevin",
                 font=font(10), bg=C['panel'], fg=C['text']).pack(pady=(4, 4))
        tk.Label(frame, text="© 2026 — Todos los derechos reservados",
                 font=font(9, italic=True), bg=C['panel'],
                 fg=C['text_l']).pack(pady=(4, 14))
        tk.Label(frame, text="Método: Newton-Raphson\nMódulos: Flujos de Potencia",
                 font=font_m(9), bg=C['panel'], fg=C['text_l'],
                 justify='center').pack()
        tk.Button(frame, text="Cerrar", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=24, pady=8, cursor='hand2',
                  borderwidth=0).pack(pady=16)

    # =================================================================
    # GUARDAR / ABRIR
    # =================================================================
    def _nuevo_proyecto(self):
        if messagebox.askyesno("Nuevo proyecto",
            "¿Crear un nuevo proyecto?\n\nSe perderán los datos actuales."):
            self.generar_tablas()
            self._set_status("Nuevo proyecto creado.", 'info')

    def _guardar_proyecto(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".vatia",
            filetypes=[("Proyecto Vatia", "*.vatia"), ("Todos", "*.*")])
        if not path: return
        try:
            n = int(self.entry_n.get())
            data = {'version': '1.0', 'n': n, 'nodos': [], 'lineas': [],
                    'base_mva': self.entry_base_mva.get(),
                    'base_kv': self.entry_base_kv.get()}
            for i in range(n):
                data['nodos'].append({
                    'tipo': self.tipo_widgets[i].get(),
                    'V': self.V_fijo_widgets[i].get(),
                    'theta': self.theta_fijo_widgets[i].get(),
                    'Pd': self.Pd_widgets[i].get(),
                    'Qd': self.Qd_widgets[i].get(),
                    'Pg': self.Pg_widgets[i].get(),
                    'Qg': self.Qg_widgets[i].get()})
            for row in self.entries_lineas:
                data['lineas'].append([e.get() for e in row])
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            self._set_status(f"Guardado: {os.path.basename(path)}", 'ok')
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))

    def _abrir_proyecto(self):
        path = filedialog.askopenfilename(
            filetypes=[("Proyecto Vatia", "*.vatia"), ("Todos", "*.*")])
        if not path: return
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)
            if 'base_mva' in data:
                self.entry_base_mva.delete(0, tk.END)
                self.entry_base_mva.insert(0, data['base_mva'])
            if 'base_kv' in data:
                self.entry_base_kv.delete(0, tk.END)
                self.entry_base_kv.insert(0, data['base_kv'])
            self._actualizar_bases()

            self.entry_n.delete(0, tk.END)
            self.entry_n.insert(0, str(data['n']))
            self.generar_tablas()
            for i, nd in enumerate(data['nodos']):
                self.tipo_widgets[i].set(nd.get('tipo', 'PQ (1)'))
                for w, key in [(self.V_fijo_widgets[i], 'V'),
                               (self.theta_fijo_widgets[i], 'theta'),
                               (self.Pd_widgets[i], 'Pd'),
                               (self.Qd_widgets[i], 'Qd'),
                               (self.Pg_widgets[i], 'Pg'),
                               (self.Qg_widgets[i], 'Qg')]:
                    w.delete(0, tk.END)
                    w.insert(0, nd.get(key, '0.0'))
                self.actualizar_fila_potencias(i)
            for i, fila in enumerate(data['lineas']):
                if i >= len(self.entries_lineas): break
                for j, val in enumerate(fila):
                    if j >= len(self.entries_lineas[i]): break
                    self.entries_lineas[i][j].delete(0, tk.END)
                    self.entries_lineas[i][j].insert(0, val)
            self._last_tipo_values = {i: self.tipo_widgets[i].get()
                                       for i in range(data['n'])}
            self._refresh_slack_options()
            self._dibujar_preview()
            self._set_status(f"Cargado: {os.path.basename(path)}", 'ok')
        except Exception as e:
            messagebox.showerror("Error al abrir", str(e))

    # =================================================================
    # EXPORTAR A EXCEL
    # =================================================================
    def exportar_excel(self):
        try:
            self._set_status("Generando Excel...", 'info')
            self.root.update_idletasks()
            n = int(self.entry_n.get())
            def _f(w, default=0.0): return self._parse_float(w.get(), default)
            datos_nodos = []; nodos_info = []
            for i in range(n):
                tipo_str = self.tipo_widgets[i].get()
                tipo = 3 if "Slack" in tipo_str else (2 if "PV" in tipo_str else 1)
                V = _f(self.V_fijo_widgets[i], 1.0)
                theta = _f(self.theta_fijo_widgets[i], 0.0)
                Pd = _f(self.Pd_widgets[i], 0.0)
                Qd = _f(self.Qd_widgets[i], 0.0)
                Pg = _f(self.Pg_widgets[i], 0.0)
                Qg = _f(self.Qg_widgets[i], 0.0)
                Pesp = (Pg - Pd) / self.BASE_MVA
                Qesp = (Qg - Qd) / self.BASE_MVA
                datos_nodos.append([i + 1, Pg, Qg, Pd, Qd,
                                     0, 0, V * 220, 220, tipo])
                nodos_info.append({'tipo': tipo, 'V': V, 'theta': theta,
                                    'Pesp': Pesp, 'Qesp': Qesp,
                                    'Pd_MW': Pd, 'Qd_MVAr': Qd,
                                    'Pg_MW': Pg, 'Qg_MVAr': Qg})
            datos_lineas = []; lineas_info = []
            for row in self.entries_lineas:
                datos_lineas.append([e.get().strip() if e.get().strip() else '0'
                                     for e in row])
                a = self._parse_int(row[0].get(), 0); b = self._parse_int(row[1].get(), 0)
                G = self._parse_float(row[2].get(), 0.0); B = self._parse_float(row[3].get(), 0.0)
                ysh = self._parse_float(row[4].get(), 0.0)
                if a > 0 and b > 0 and a <= n and b <= n and a != b:
                    lineas_info.append({'a': a - 1, 'b': b - 1,
                                        'G': G, 'B': B, 'y_shunt': ysh})
            crear_excel_completo(n, datos_nodos, datos_lineas,
                                  dict(self.datos_iteracion), nodos_info, lineas_info)
            self._set_status("Excel exportado correctamente.", 'ok')
        except Exception as e:
            import traceback
            messagebox.showerror("Error",
                f"Error al exportar:\n{e}\n\n{traceback.format_exc()}")

    # =================================================================
    # NAVEGACIÓN Y CIERRE
    # =================================================================
    def _volver_al_menu(self):
        if not messagebox.askyesno("Confirmar",
            "¿Volver al menú principal?\nSe perderán los datos no guardados."):
            return
        self._detener_animaciones()
        for w in self.root.winfo_children():
            w.destroy()
        try:
            def abrir_simulador(root):
                root.title("Vatia — Flujos de Potencia (Newton-Raphson)")
                root.geometry("1500x900"); root.minsize(1180, 720)
                root.configure(bg="#EDF1F6")
                app = AppFlujoPotencia(root); root.app = app
            DashboardScreen(self.root, abrir_simulador)
        except Exception:
            import traceback
            traceback.print_exc()

    def _on_close(self):
        self._detener_animaciones()
        try:
            self.root.destroy()
        except Exception:
            pass
# =====================================================================
# DASHBOARD DE MÓDULOS
# =====================================================================
class DashboardScreen:
    """Pantalla de selección de módulos con logos animados y teoría."""

    C = {
        'bg':          '#EDF1F6',
        'panel':       '#FFFFFF',
        'surface':     '#F7F9FC',
        'primary':     '#1E3A5F',
        'primary_l':   '#2C5282',
        'accent':      '#0E7490',
        'accent_l':    '#0891B2',
        'text':        '#1F2937',
        'text_l':      '#4B5563',
        'text_ll':     '#94A3B8',
        'border':      '#CBD5E1',
        'slack':       '#047857',
        'pv':          '#6D28D9',
        'pq':          '#C2410C',
        'inactive':    '#94A3B8',
        'inactive_bg': '#F1F5F9',
    }

    def __init__(self, root, on_select_power_flow):
        self.root = root
        self.on_select_power_flow = on_select_power_flow
        C = self.C

        self.root.title("Vatia — Suite de Análisis de Sistemas de Potencia")
        self.root.geometry("1340x860")
        self.root.minsize(1150, 760)
        self.root.configure(bg=C['bg'])

        # Estado de animación
        self._dash_logo_pulse = 0.0
        self._dash_flow_particles = []
        self._dash_flow_counter = 0
        self._dash_logo_cx = 40
        self._dash_logo_cy = 40
        self._dash_logo_size = 68
        self._logo_animation_items = []
        self._logo_animation_id = None
        self._dash_logo_after_id = None

        # ========= HEADER =========
        hdr = tk.Frame(self.root, bg=C['primary'], height=110)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['accent'], height=4).pack(fill=tk.X, side=tk.TOP)

        centro = tk.Frame(hdr, bg=C['primary'])
        centro.pack(expand=True, pady=10)
        self._dash_logo_canvas = tk.Canvas(centro, width=440, height=80,
                                             bg=C['primary'],
                                             highlightthickness=0)
        self._dash_logo_canvas.pack()

        right = tk.Frame(hdr, bg=C['primary'])
        right.place(relx=1.0, rely=0.5, anchor='e', x=-32)
        tk.Label(right, text="por Kevin", font=font(11, italic=True),
                 bg=C['primary'], fg='#CBD5E1').pack(anchor='e')
        tk.Label(right, text="v1.0  ·  2026", font=font_m(9),
                 bg=C['primary'], fg='#94A3B8').pack(anchor='e')

        # ========= SUBTÍTULO =========
        subtitle = tk.Frame(self.root, bg=C['bg'])
        subtitle.pack(fill=tk.X, pady=(32, 12))
        tk.Label(subtitle, text="Seleccione un módulo para comenzar",
                 font=font(20, "bold"), bg=C['bg'], fg=C['text']).pack()
        tk.Label(subtitle,
                 text="Suite integral de análisis para sistemas eléctricos de potencia",
                 font=font(11, italic=True), bg=C['bg'],
                 fg=C['text_l']).pack(pady=(6, 0))

        # ========= TARJETAS =========
        cards_frame = tk.Frame(self.root, bg=C['bg'])
        cards_frame.pack(expand=True, fill=tk.BOTH, padx=50, pady=20)
        for i in range(3):
            cards_frame.grid_columnconfigure(i, weight=1, uniform='col')
        cards_frame.grid_rowconfigure(0, weight=1)

        self._crear_tarjeta(
            cards_frame, col=0,
            logo_fn=self._dibujar_logo_flujos,
            logotipo='flujos',
            titulo="Flujos de Potencia",
            subtitulo="Newton-Raphson",
            descripcion="Análisis de flujo de carga iterativo. Cálculo de "
                        "tensiones nodales, potencias y pérdidas con "
                        "documentación paso a paso de cada iteración.",
            color=C['accent'],
            estado='ACTIVO',
            accion=self._seleccionar_flujos)

        self._crear_tarjeta(
            cards_frame, col=1,
            logo_fn=self._dibujar_logo_fallas,
            logotipo='fallas',
            titulo="Análisis de Fallas",
            subtitulo="Componentes simétricas",
            descripcion="Cálculo de corrientes de falla trifásica, "
                        "monofásica y bifásica. Impedancias de Thévenin, "
                        "reconstrucción abc y aportes por generador.",
            color=C['pq'],
            estado='ACTIVO',
            accion=self._abrir_modulo_fallas)

        self._crear_tarjeta(
            cards_frame, col=2,
            logo_fn=self._dibujar_logo_despacho,
            logotipo='despacho',
            titulo="Despacho Económico",
            subtitulo="Optimización + congestión",
            descripcion="Despacho económico con análisis de congestión, "
                        "redespacho óptimo y precios marginales nodales.",
            color=C['pv'],
            estado='ACTIVO',
            accion=self._abrir_modulo_despacho)

        # ========= FOOTER =========
        footer = tk.Frame(self.root, bg=C['primary'], height=38)
        footer.pack(fill=tk.X, side=tk.BOTTOM)
        footer.pack_propagate(False)
        tk.Label(footer,
                 text="© 2026 Vatia · Suite académica de Sistemas de Potencia  ·  por Kevin",
                 font=font(9), bg=C['primary'], fg='#CBD5E1').pack(pady=10)

        # Arrancar animaciones
        self.root.after(150, self._animate_dash_logo)
        self.root.after(150, self._animar_logos_modulos)

    # ----------------------------------------------------------------
    # LOGOS POR MÓDULO (animados)
    # ----------------------------------------------------------------
    def _dibujar_logo_flujos(self, canvas, cx, cy, size, color, phase=0.0):
        canvas.delete("all")
        C = self.C
        r = size / 2

        pts = []
        for i in range(6):
            ang = math.pi * i / 3 - math.pi / 2
            pts.extend([cx + r * math.cos(ang), cy + r * math.sin(ang)])
        canvas.create_polygon(pts, fill=color, outline='')

        colores = [C['slack'], C['pv'], C['pq'], C['slack']]
        for i in range(4):
            ang = 2 * math.pi * i / 4 - math.pi / 2
            radio = r * 0.7 + 3 * math.sin(phase * 2 * math.pi + i * 1.2)
            nx = cx + radio * math.cos(ang)
            ny = cy + radio * math.sin(ang)

            pulse = 0.5 + 0.5 * abs(math.sin(phase * 2 * math.pi + i * 0.5))
            halo_r = 10 + 4 * pulse
            canvas.create_oval(nx - halo_r, ny - halo_r,
                                nx + halo_r, ny + halo_r,
                                fill='', outline='white', width=1.5)

            canvas.create_line(cx, cy, nx, ny, fill='white', width=2)
            canvas.create_oval(nx - 5, ny - 5, nx + 5, ny + 5,
                                fill=colores[i], outline='white', width=2)

        brillo = 0.6 + 0.4 * abs(math.sin(phase * 2 * math.pi))
        v_h = r * 1.0
        v_w = r * 0.5
        canvas.create_line(cx - v_w, cy - v_h/2, cx, cy + v_h/2,
                            cx + v_w, cy - v_h/2,
                            fill='white',
                            width=4 + 2 * brillo,
                            capstyle=tk.ROUND, joinstyle=tk.ROUND)

    def _dibujar_logo_fallas(self, canvas, cx, cy, size, color, phase=0.0):
        canvas.delete("all")
        r = size / 2
        canvas.create_oval(cx - r, cy - r, cx + r, cy + r,
                            fill=color, outline='')

        pts = [
            cx + r*0.15, cy - r*0.75,
            cx - r*0.35, cy + r*0.05,
            cx - r*0.05, cy + r*0.05,
            cx - r*0.15, cy + r*0.75,
            cx + r*0.35, cy - r*0.05,
            cx + r*0.05, cy - r*0.05,
        ]
        canvas.create_polygon(pts, fill='white', outline='')

        for k in range(4):
            ang = phase * 2 * math.pi + k * math.pi / 2
            dx = r * 0.75 * math.cos(ang)
            dy = r * 0.75 * math.sin(ang)
            destello_r = 3 + 3 * abs(math.sin(phase * 4 * math.pi + k))
            canvas.create_oval(cx + dx - destello_r, cy + dy - destello_r,
                                cx + dx + destello_r, cy + dy + destello_r,
                                fill='white', outline='')

    def _dibujar_logo_despacho(self, canvas, cx, cy, size, color, phase=0.0):
        canvas.delete("all")
        r = size / 2
        canvas.create_oval(cx - r, cy - r, cx + r, cy + r,
                            fill=color, outline='')

        bar_w = r * 0.18
        alturas_base = [0.5, 0.75, 0.95]
        for i in range(3):
            fase_barra = (phase + i * 0.15) % 1.0
            pulso = 0.85 + 0.15 * math.sin(fase_barra * 2 * math.pi)
            h = r * alturas_base[i] * pulso
            x0 = cx - r*0.65 + i * (bar_w + r*0.1)
            y0 = cy + r*0.5
            canvas.create_rectangle(x0, y0 - h, x0 + bar_w, y0,
                                     fill='white', outline='')

        canvas.create_line(cx - r*0.75, cy + r*0.5,
                            cx + r*0.75, cy + r*0.5,
                            fill='white', width=2)

        offset = 3 * math.sin(phase * 2 * math.pi)
        canvas.create_line(cx + r*0.4, cy - r*0.6 + offset,
                            cx + r*0.7, cy - r*0.75 + offset,
                            fill='white', width=2)
        canvas.create_polygon(
            cx + r*0.75, cy - r*0.78 + offset,
            cx + r*0.62, cy - r*0.65 + offset,
            cx + r*0.75, cy - r*0.60 + offset,
            fill='white', outline='')

    # ----------------------------------------------------------------
    # TARJETA
    # ----------------------------------------------------------------
    def _crear_tarjeta(self, parent, col, logo_fn, logotipo, titulo, subtitulo,
                       descripcion, color, estado, accion):
        C = self.C
        activo = (estado == 'ACTIVO')

        outer = tk.Frame(parent, bg=C['bg'])
        outer.grid(row=0, column=col, sticky='nsew', padx=10, pady=10)

        shadow = tk.Frame(outer, bg='#D5DCE3')
        shadow.pack(fill=tk.BOTH, expand=True)

        card = tk.Frame(shadow, bg=C['panel'],
                        highlightthickness=2,
                        highlightbackground=color if activo else C['border'])
        card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        tk.Frame(card, bg=color if activo else C['inactive'],
                 height=6).pack(fill=tk.X)

        logo_canvas = tk.Canvas(card, width=120, height=120,
                                  bg=C['panel'], highlightthickness=0)
        logo_canvas.pack(pady=(24, 10))
        self._logo_animation_items.append({
            'canvas': logo_canvas,
            'fn': logo_fn,
            'cx': 60, 'cy': 60, 'size': 110,
            'color': color if activo else C['inactive'],
        })

        tk.Label(card, text=titulo, font=font(15, "bold"),
                 bg=C['panel'],
                 fg=C['text'] if activo else C['inactive']).pack()

        tk.Label(card, text=subtitulo, font=font(10, italic=True),
                 bg=C['panel'], fg=C['text_l']).pack(pady=(2, 10))

        badge_bg = color if activo else C['inactive_bg']
        badge_fg = 'white' if activo else C['inactive']
        badge = tk.Frame(card, bg=badge_bg)
        badge.pack()
        tk.Label(badge, text=f"  {estado}  ",
                 font=font(8, "bold"),
                 bg=badge_bg, fg=badge_fg,
                 padx=8, pady=4).pack()

        tk.Label(card, text=descripcion, font=font(10),
                 bg=C['panel'], fg=C['text_l'],
                 wraplength=280, justify='center').pack(padx=20, pady=(14, 16))

        btn_text = "Abrir módulo" if activo else "No disponible"
        btn = tk.Button(card, text=btn_text,
                        command=accion if activo else lambda: None,
                        bg=color if activo else C['inactive_bg'],
                        fg='white' if activo else C['inactive'],
                        font=font(11, "bold"),
                        relief=tk.FLAT,
                        cursor='hand2' if activo else 'arrow',
                        state=tk.NORMAL if activo else tk.DISABLED,
                        padx=20, pady=11, borderwidth=0,
                        activebackground=C['primary_l'] if activo else C['inactive_bg'],
                        activeforeground='white',
                        disabledforeground=C['inactive'])
        btn.pack(fill=tk.X, padx=30, pady=(0, 8))

        btn_teoria = tk.Button(card, text="Ver teoría",
                                command=lambda t=logotipo: self._mostrar_teoria(t),
                                bg=C['panel'], fg=color if activo else C['text_l'],
                                font=font(10, "bold"),
                                relief=tk.SOLID, borderwidth=1,
                                cursor='hand2',
                                padx=20, pady=9,
                                activebackground=C['bg'],
                                activeforeground=color if activo else C['text'])
        btn_teoria.pack(fill=tk.X, padx=30, pady=(0, 24))
        btn_teoria.config(highlightbackground=color if activo else C['border'],
                          highlightthickness=1)

        if activo:
            def on_enter(e):
                card.config(highlightbackground=C['primary'])
                btn.config(bg=C['primary_l'])
            def on_leave(e):
                card.config(highlightbackground=color)
                btn.config(bg=color)
            for w in (card, btn):
                w.bind('<Enter>', on_enter)
                w.bind('<Leave>', on_leave)

    # ----------------------------------------------------------------
    # ANIMACIÓN DE LOGOS DE MÓDULOS
    # ----------------------------------------------------------------
    def _animar_logos_modulos(self):
        try:
            for item in self._logo_animation_items:
                canvas = item['canvas']
                try:
                    if not canvas.winfo_exists():
                        continue
                except Exception:
                    continue
                item['fn'](canvas, item['cx'], item['cy'],
                            item['size'], item['color'],
                            phase=self._dash_logo_pulse)
            self._logo_animation_id = self.root.after(40, self._animar_logos_modulos)
        except Exception:
            pass

    # ----------------------------------------------------------------
    # LOGO ANIMADO DEL HEADER
    # ----------------------------------------------------------------
    def _animate_dash_logo(self):
        if not hasattr(self, '_dash_logo_canvas'):
            return
        try:
            if not self._dash_logo_canvas.winfo_exists():
                return
        except Exception:
            return
        try:
            C = self.C
            self._dash_logo_pulse = (self._dash_logo_pulse + 0.025) % 1.0
            t = self._dash_logo_pulse
            v_brightness = 0.35 + 0.65 * abs(math.sin(t * 2 * math.pi))
            v_color = interp_color('#0E7490', '#FFFFFF', v_brightness)
            pulse_nodes = 0.5 + 0.5 * abs(math.sin(t * 2 * math.pi))

            self._dash_logo_canvas.delete("all")
            self._draw_dash_logo(self._dash_logo_cx, self._dash_logo_cy,
                                   self._dash_logo_size, v_color, pulse_nodes, t)

            self._dash_logo_canvas.create_text(
                self._dash_logo_cx + 58, self._dash_logo_cy - 8,
                text="V A T I A",
                font=font(22, "bold", family=FONT_FAMILY_ALT),
                fill='#FFFFFF', anchor='w')
            self._dash_logo_canvas.create_text(
                self._dash_logo_cx + 60, self._dash_logo_cy + 16,
                text="Suite Profesional de Análisis de Sistemas de Potencia",
                font=font(8, italic=True), fill='#CBD5E1', anchor='w')

            self._dash_flow_counter += 1
            if self._dash_flow_counter % 8 == 0:
                self._dash_spawn_flow()
            self._dash_draw_flows()
            self._dash_logo_after_id = self.root.after(40, self._animate_dash_logo)
        except tk.TclError:
            return
        except Exception:
            return

    def _draw_dash_logo(self, cx, cy, size, v_color, pulse_nodes, t):
        C = self.C
        r = size / 2
        lw_main = max(2, int(size / 14))
        lw_thin = max(1, int(size / 26))
        hex_pts = []
        for i in range(6):
            ang = 2 * math.pi * i / 6 - math.pi / 2
            hex_pts.extend([cx + r * math.cos(ang), cy + r * math.sin(ang)])
        self._dash_logo_canvas.create_polygon(hex_pts, fill='#152741', outline='')
        tipos = [C['slack'], C['pv'], C['pq'], C['pq'], C['pv'], C['slack']]
        nodo_r = size * 0.065
        self._dash_node_positions = []
        for i in range(6):
            ang = 2 * math.pi * i / 6 - math.pi / 2
            nx = cx + r * 0.88 * math.cos(ang)
            ny = cy + r * 0.88 * math.sin(ang)
            self._dash_node_positions.append((nx, ny, i))
            self._dash_logo_canvas.create_line(cx, cy, nx, ny,
                                                 fill=C['accent'], width=lw_thin)
            phase_node = (t + i * 0.15) % 1.0
            nodo_pulse = 0.4 + 0.6 * abs(math.sin(phase_node * 2 * math.pi))
            halo_r = nodo_r * (1.2 + 0.5 * nodo_pulse)
            self._dash_logo_canvas.create_oval(nx - halo_r, ny - halo_r,
                                                 nx + halo_r, ny + halo_r,
                                                 fill=C['accent'], outline='')
            self._dash_logo_canvas.create_oval(nx - nodo_r, ny - nodo_r,
                                                 nx + nodo_r, ny + nodo_r,
                                                 fill=tipos[i], outline='#FFFFFF',
                                                 width=max(1, int(size / 40)))
        v_h = r * 1.1
        v_w = r * 0.6
        pts = [cx - v_w, cy - v_h / 2, cx, cy + v_h / 2, cx + v_w, cy - v_h / 2]
        self._dash_logo_canvas.create_line(*pts, fill='#0A1420',
                                             width=lw_main + 6,
                                             capstyle=tk.ROUND,
                                             joinstyle=tk.ROUND)
        self._dash_logo_canvas.create_line(*pts, fill=v_color,
                                             width=lw_main + 2,
                                             capstyle=tk.ROUND,
                                             joinstyle=tk.ROUND)

    def _dash_spawn_flow(self):
        import random
        if not hasattr(self, '_dash_node_positions'): return
        if len(self._dash_flow_particles) > 12: return
        idx = random.randint(0, 5)
        nx, ny, _ = self._dash_node_positions[idx]
        self._dash_flow_particles.append({
            'x': nx, 'y': ny,
            'tx': self._dash_logo_cx, 'ty': self._dash_logo_cy,
            'progress': 0.0,
            'speed': 0.04 + random.random() * 0.02})

    def _dash_draw_flows(self):
        activas = []
        for p in self._dash_flow_particles:
            p['progress'] += p['speed']
            if p['progress'] >= 1.0:
                continue
            x = p['x'] + (p['tx'] - p['x']) * p['progress']
            y = p['y'] + (p['ty'] - p['y']) * p['progress']
            rad = 3 - 2 * p['progress']
            bright = 0.5 + 0.5 * p['progress']
            color = interp_color('#0E7490', '#FFFFFF', bright)
            self._dash_logo_canvas.create_oval(x - rad, y - rad,
                                                 x + rad, y + rad,
                                                 fill=color, outline='')
            activas.append(p)
        self._dash_flow_particles = activas

    # ----------------------------------------------------------------
    # TEORÍA POR MÓDULO
    # ----------------------------------------------------------------
    def _mostrar_teoria(self, tipo):
        if tipo == 'flujos':
            self._teoria_flujos()
        elif tipo == 'fallas':
            # Abrir el módulo y mostrar la pestaña de teoría dentro
            self._abrir_modulo_fallas()
            # (si quieres, luego auto-abrir la teoría)
        elif tipo == 'despacho':
            self._abrir_modulo_despacho()
                

    def _teoria_placeholder(self, titulo, mensaje):
        C = self.C
        win = tk.Toplevel(self.root)
        win.title(f"Teoría — {titulo}")
        win.geometry("560x320")
        win.configure(bg=C['bg'])
        win.transient(self.root)

        hdr = tk.Frame(win, bg=C['primary'], height=64)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text=f"  {titulo}", font=font(14, "bold"),
                 bg=C['primary'], fg='white').pack(side=tk.LEFT, padx=20, pady=16)
        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['pq'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=14, pady=6, cursor='hand2',
                  borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=14)

        body = tk.Frame(win, bg=C['panel'])
        body.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        tk.Label(body, text="Módulo en desarrollo",
                 font=font(13, "bold"), bg=C['panel'], fg=C['pq']
                 ).pack(pady=(20, 12))
        tk.Label(body, text=mensaje, font=font(11), bg=C['panel'],
                 fg=C['text_l'], wraplength=480, justify='center'
                 ).pack(padx=20, pady=10)

        tk.Button(body, text="Cerrar", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=24, pady=8, cursor='hand2',
                  borderwidth=0).pack(pady=20)

    def _teoria_flujos(self):
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Teoría — Flujos de Potencia (Newton-Raphson)")
        win.geometry("1000x760")
        win.configure(bg=C['bg'])
        win.transient(self.root)

        hdr = tk.Frame(win, bg=C['primary'], height=70)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  Teoría — Flujos de Potencia con Newton-Raphson",
                 font=font(15, "bold"), bg=C['primary'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=18)
        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['pq'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=14, pady=6, cursor='hand2',
                  borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=14)

        canvas = tk.Canvas(win, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(win, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas.find_all()[0],
                                                 width=e.width))
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        def _wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        canvas.bind('<MouseWheel>', _wheel)
        inner.bind('<MouseWheel>', _wheel)

        def titulo_seccion(txt, subtitulo=''):
            f = tk.Frame(inner, bg=C['primary'])
            f.pack(fill=tk.X, padx=30, pady=(24, 10))
            tk.Label(f, text=txt, font=font(14, "bold"),
                     bg=C['primary'], fg='white', anchor='w',
                     padx=16, pady=10).pack(fill=tk.X)
            if subtitulo:
                tk.Label(inner, text=subtitulo, font=font(10, italic=True),
                         bg=C['bg'], fg=C['text_l'], anchor='w'
                         ).pack(fill=tk.X, padx=34, pady=(0, 8))

        def parrafo(txt, bold=False):
            tk.Label(inner, text=txt,
                     font=font(11, "bold" if bold else "normal"),
                     bg=C['bg'], fg=C['text'],
                     justify='left', anchor='w', wraplength=850
                     ).pack(fill=tk.X, padx=34, pady=3)

        def formula(txt):
            f = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                         highlightbackground=C['border'])
            f.pack(fill=tk.X, padx=50, pady=6)
            tk.Label(f, text=txt, font=font_m(11), bg=C['surface'],
                     fg=C['primary'], justify='left', anchor='w',
                     padx=16, pady=10).pack(fill=tk.X)

        def paso(numero, titulo, cuerpo, formulas=None):
            head = tk.Frame(inner, bg=C['panel'], highlightthickness=1,
                            highlightbackground=C['border'])
            head.pack(fill=tk.X, padx=30, pady=(16, 0))
            tk.Frame(head, bg=C['accent'], width=6).pack(side=tk.LEFT, fill=tk.Y)
            num_box = tk.Frame(head, bg=C['accent'], width=46, height=46)
            num_box.pack(side=tk.LEFT, padx=12, pady=12)
            num_box.pack_propagate(False)
            tk.Label(num_box, text=str(numero), font=font(20, "bold"),
                     bg=C['accent'], fg='white').pack(expand=True)
            tk.Label(head, text=titulo, font=font(13, "bold"),
                     bg=C['panel'], fg=C['primary'], padx=8, pady=16
                     ).pack(side=tk.LEFT)
            parrafo(cuerpo)
            if formulas:
                for f in formulas:
                    formula(f)

        titulo_seccion("Método de Newton-Raphson para Flujos de Potencia",
                        "Resolución iterativa del sistema de ecuaciones no lineales de potencia nodal")

        parrafo(
            "El método de Newton-Raphson permite resolver el sistema de "
            "ecuaciones no lineales que describen el comportamiento de una red "
            "eléctrica de potencia. A diferencia de otros métodos iterativos "
            "(como Gauss-Seidel), converge cuadráticamente: cada iteración "
            "duplica aproximadamente el número de dígitos correctos.")
        parrafo(
            "El sistema se describe mediante dos ecuaciones por nodo (potencia "
            "activa P y reactiva Q), en función de las tensiones nodales V y "
            "los ángulos θ. El objetivo es encontrar los valores de V y θ que "
            "satisfacen todas las ecuaciones simultáneamente.")

        titulo_seccion("Clasificación de Nodos")
        parrafo("Antes de aplicar el método, los nodos se clasifican en tres tipos según las variables conocidas:")
        parrafo("• Slack (Barra Infinita): V y θ conocidos. P y Q se calculan.")
        parrafo("• PV (Generador): P y |V| conocidos. Q y θ se calculan.")
        parrafo("• PQ (Carga): P y Q conocidos. |V| y θ se calculan.")

        titulo_seccion("Los 7 Pasos del Método Newton-Raphson")

        paso(1, "Inicialización",
             "Se asignan valores iniciales a las tensiones y ángulos de todos los "
             "nodos. Típicamente se usa el 'flat start': V = 1.0 p.u. y θ = 0° "
             "para todos los nodos PQ y PV. Los nodos Slack mantienen sus valores "
             "fijos. También se construye la matriz de admitancias nodales (Ybus) "
             "a partir de los datos de líneas.",
             ["V_i⁽⁰⁾ = 1.0 p.u.     θ_i⁽⁰⁾ = 0 rad     (para nodos PQ y PV)",
              "Y_bus = matriz de admitancias nodales (n × n)"])

        paso(2, "Cálculo de Potencias Calculadas",
             "Con los valores actuales de V y θ, se calculan las potencias activa "
             "y reactiva inyectadas en cada nodo usando las ecuaciones de "
             "potencia nodal. Estas son funciones no lineales de las tensiones "
             "y los ángulos.",
             ["P_i = V_i · Σⱼ V_j (G_ij·cos θ_ij + B_ij·sin θ_ij)",
              "Q_i = V_i · Σⱼ V_j (G_ij·sin θ_ij − B_ij·cos θ_ij)",
              "donde  θ_ij = θ_i − θ_j"])

        paso(3, "Cálculo del Mismatch (ΔP, ΔQ)",
             "Se calcula la diferencia entre la potencia especificada (dato del "
             "problema) y la potencia calculada en el paso anterior. Este "
             "'mismatch' indica cuánto se desvía el sistema de la solución.",
             ["ΔP_i = P_esp,i − P_calc,i    (para nodos PQ y PV)",
              "ΔQ_i = Q_esp,i − Q_calc,i    (solo para nodos PQ)"])

        paso(4, "Verificación de Convergencia",
             "Se verifica si el mismatch está dentro de la tolerancia establecida "
             "(por ejemplo 1×10⁻³ p.u.). Si max|ΔP| y max|ΔQ| son menores que la "
             "tolerancia, el método ha convergido y termina. Si no, se continúa "
             "con el siguiente paso.",
             ["Si  max|ΔP|, max|ΔQ| < ε   →   CONVERGIÓ  →  FIN",
              "Si no                          →   continuar al paso 5"])

        paso(5, "Construcción del Jacobiano",
             "Se construye la matriz Jacobiana, que contiene las derivadas "
             "parciales de las potencias P y Q respecto a las incógnitas θ y V. "
             "El Jacobiano se organiza en cuatro bloques: H = ∂P/∂θ, N = ∂P/∂V, "
             "M = ∂Q/∂θ, L = ∂Q/∂V.",
             ["        ┌         ┐",
              "    J = │ H    N  │   H = ∂P/∂θ    N = ∂P/∂V",
              "        │ M    L  │   M = ∂Q/∂θ    L = ∂Q/∂V",
              "        └         ┘"])

        paso(6, "Resolución del Sistema Lineal",
             "Se resuelve el sistema lineal J · Δx = mismatch, donde Δx contiene "
             "los incrementos de las incógnitas (Δθ y ΔV). Esto da la dirección "
             "y magnitud del ajuste que hay que hacer en las variables de estado.",
             ["J · Δx = mismatch",
              "Δx = J⁻¹ · mismatch",
              "donde  Δx = [Δθ₁ … Δθₙ  ΔV₁ … ΔVₚ]ᵀ",
              "       mismatch = [ΔP₁ … ΔPₙ  ΔQ₁ … ΔQₚ]ᵀ"])

        paso(7, "Actualización de Variables",
             "Se actualizan los valores de las tensiones y los ángulos de los "
             "nodos PQ y PV sumando los incrementos calculados. Luego se vuelve "
             "al paso 2 y se repite el proceso hasta alcanzar la convergencia.",
             ["θ_i⁽ᵏ⁺¹⁾ = θ_i⁽ᵏ⁾ + Δθ_i⁽ᵏ⁾     (nodos PQ y PV)",
              "V_i⁽ᵏ⁺¹⁾ = V_i⁽ᵏ⁾ + ΔV_i⁽ᵏ⁾     (nodos PQ)",
              "→  Volver al paso 2"])

        titulo_seccion("Convergencia Cuadrática")
        parrafo(
            "La gran ventaja del método Newton-Raphson es su convergencia "
            "cuadrática cerca de la solución. Esto significa que si el error "
            "en una iteración es ε, en la siguiente será del orden de ε². "
            "En la práctica, con 3 a 5 iteraciones se alcanza precisión de "
            "4 a 6 cifras decimales.")

        tk.Button(inner, text="Cerrar",
                  command=win.destroy,
                  bg=C['primary'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=32, pady=10, cursor='hand2',
                  borderwidth=0).pack(pady=30)

    def _seleccionar_flujos(self):
        """Modal para elegir entre NR completo y NR lineal (DC)."""
        print("[DEBUG] _seleccionar_flujos llamado")

        C = self.C
        modal = tk.Toplevel(self.root)
        modal.title("Selección de método")
        modal.geometry("820x500")
        modal.configure(bg=C['bg'])

        # === BLINDAJE: garantizar que se vea arriba ===
        try:
            modal.transient(self.root)
        except Exception:
            pass
        try:
            modal.attributes('-topmost', True)
            modal.after(100, lambda: modal.attributes('-topmost', False))
        except Exception:
            pass
        try:
            modal.grab_set()
        except Exception:
            pass

        # Centrar
        self.root.update_idletasks()
        try:
            x = self.root.winfo_x() + (self.root.winfo_width() - 820) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 500) // 2
            modal.geometry(f"+{max(0, x)}+{max(0, y)}")
        except Exception:
            pass

        modal.lift()
        modal.focus_force()
        print("[DEBUG] Modal creado y posicionado")

        # === Header ===
        hdr = tk.Frame(modal, bg=C['primary'], height=70)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['accent'], height=4).pack(fill=tk.X, side=tk.TOP)
        tk.Label(hdr, text="Seleccione el método de resolución",
                 font=font(15, "bold"), bg=C['primary'], fg='white'
                 ).pack(pady=20)

        # === Cuerpo ===
        frame = tk.Frame(modal, bg=C['bg'])
        frame.pack(expand=True, fill=tk.BOTH, padx=24, pady=20)

        # Cada card en un frame separado, posicionados con pack lado a lado
        contenedor = tk.Frame(frame, bg=C['bg'])
        contenedor.pack(expand=True, fill=tk.BOTH)

        def _crear_opcion(parent, icono_txt, titulo, subtitulo,
                          descripcion, color, comando):
            outer = tk.Frame(parent, bg=C['bg'])
            outer.pack(side=tk.LEFT, expand=True, fill=tk.BOTH,
                       padx=10, pady=10)

            card = tk.Frame(outer, bg=C['panel'],
                            highlightthickness=2, highlightbackground=color)
            card.pack(fill=tk.BOTH, expand=True)

            tk.Frame(card, bg=color, height=5).pack(fill=tk.X)

            cv = tk.Canvas(card, width=80, height=80, bg=C['panel'],
                            highlightthickness=0)
            cv.pack(pady=(22, 12))
            cv.create_oval(4, 4, 76, 76, fill=color, outline='')
            cv.create_text(40, 40, text=icono_txt,
                            font=font(22, "bold"), fill='white')

            tk.Label(card, text=titulo, font=font(14, "bold"),
                     bg=C['panel'], fg=C['text']).pack()
            tk.Label(card, text=subtitulo, font=font(9, italic=True),
                     bg=C['panel'], fg=C['text_l']).pack(pady=(2, 8))
            tk.Label(card, text=descripcion, font=font(9),
                     bg=C['panel'], fg=C['text_l'],
                     wraplength=280, justify='center'
                     ).pack(padx=16, pady=(0, 16))

            btn = tk.Button(card, text="Abrir método", command=comando,
                            bg=color, fg='white', font=font(11, "bold"),
                            relief=tk.FLAT, cursor='hand2',
                            padx=20, pady=11, borderwidth=0,
                            activebackground=color, activeforeground='white')
            btn.pack(fill=tk.X, padx=24, pady=(0, 22))

        def cerrar_modal():
            try:
                print("[DEBUG] Cerrando modal")
                modal.grab_release()
            except Exception:
                pass
            try:
                modal.destroy()
            except Exception:
                pass

        def _detener_animaciones():
            try:
                if hasattr(self, '_logo_animation_id') and self._logo_animation_id:
                    self.root.after_cancel(self._logo_animation_id)
                    self._logo_animation_id = None
                if hasattr(self, '_dash_logo_after_id') and self._dash_logo_after_id:
                    self.root.after_cancel(self._dash_logo_after_id)
                    self._dash_logo_after_id = None
            except Exception:
                pass

        def abrir_completo():
            print("[DEBUG] Abriendo NR completo")
            cerrar_modal()
            _detener_animaciones()
            for w in self.root.winfo_children():
                w.destroy()
            self.on_select_power_flow(self.root)

        def abrir_lineal():
            print("[DEBUG] Abriendo NR lineal (DC)")
            cerrar_modal()
            _detener_animaciones()
            for w in self.root.winfo_children():
                w.destroy()
            try:
                self.root.title("Vatia — Flujos de Potencia Lineal (DC)")
                self.root.geometry("1500x900")
                self.root.minsize(1180, 720)
                self.root.configure(bg="#EDF1F6")
                app = LinearFlowScreen(self.root)
                self.root.app = app
                print("[DEBUG] LinearFlowScreen abierto")
            except Exception as e:
                import traceback
                traceback.print_exc()
                messagebox.showerror(
                    "Error",
                    f"Error al abrir el módulo DC:\n\n{traceback.format_exc()}")

        _crear_opcion(contenedor, "NR",
                       "Newton-Raphson\nCompleto",
                       "Método iterativo exacto",
                       "Resuelve el sistema no lineal de ecuaciones de "
                       "potencia usando iteraciones. Preciso y detallado.",
                       C['accent'], abrir_completo)

        _crear_opcion(contenedor, "DC",
                       "Newton-Raphson\nLineal (DC)",
                       "Aproximación lineal",
                       "Aproximación lineal que asume redes inductivas y "
                       "tensiones cercanas a 1.0 p.u. Rápido y sin iteraciones.",
                       '#B45309', abrir_lineal)

        # === Botón cancelar (dentro del modal, no del contenedor) ===
        tk.Button(modal, text="Cancelar", command=cerrar_modal,
                  bg=C['surface'], fg=C['text'], font=font(10, "bold"),
                  relief=tk.FLAT, padx=24, pady=8, cursor='hand2',
                  borderwidth=0).pack(pady=(0, 16))

        # Al cerrar con X
        modal.protocol("WM_DELETE_WINDOW", cerrar_modal)

        print("[DEBUG] Modal completo")

    def _abrir_modulo_fallas(self):
        try:
            for w in self.root.winfo_children():
                w.destroy()
            self.root.title("Vatia — Análisis de Fallas")
            self.root.geometry("1500x900")
            self.root.minsize(1180, 720)
            self.root.configure(bg="#EDF1F6")
            app = FaultAnalysisScreen(self.root)
            self.root.app = app
        except Exception:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error",
                f"No se pudo abrir el módulo:\n\n{traceback.format_exc()}")

    def _abrir_modulo_despacho(self):
        try:
            for w in self.root.winfo_children():
                w.destroy()
            self.root.title("Vatia — Despacho Económico")
            self.root.geometry("1500x900")
            self.root.configure(bg="#EDF1F6")
            app = EconomicDispatchScreen(self.root)
            self.root.app = app
        except Exception:
            import traceback; traceback.print_exc()
            messagebox.showerror("Error",
                f"No se pudo abrir el módulo:\n\n{traceback.format_exc()}")
# =====================================================================
# MÓDULO — FLUJO DC LINEAL (con sidebar + preview)
# =====================================================================
class LinearFlowScreen:
    """Módulo de flujo de potencia lineal (DC) con UI completa."""

    C = {
        'bg':          '#EDF1F6',
        'panel':       '#FFFFFF',
        'surface':     '#F7F9FC',
        'sidebar':     '#E2E8F0',
        'primary':     '#1E3A5F',
        'primary_l':   '#2C5282',
        'primary_d':   '#152741',
        'accent':      '#0E7490',
        'accent_l':    '#0891B2',
        'slack':       '#047857',
        'pv':          '#6D28D9',
        'pq':          '#C2410C',
        'info':        '#0369A1',
        'amber':       '#B45309',
        'text':        '#1F2937',
        'text_l':      '#4B5563',
        'text_ll':     '#94A3B8',
        'border':      '#CBD5E1',
        'head_bg':     '#E2E8F0',
        'row_even':    '#FFFFFF',
        'row_odd':     '#F7F9FC',
        'danger':      '#B91C1C',
        'success':     '#047857',
        'gray_dark':   '#475569',
    }
    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # === ARCHIVO ===
        m_arch = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Archivo", menu=m_arch)
        m_arch.add_command(label="Nuevo proyecto\t\tCtrl+N",
                            command=self._nuevo_proyecto)
        m_arch.add_command(label="Abrir proyecto...\tCtrl+O",
                            command=self._abrir_proyecto)
        m_arch.add_command(label="Guardar\t\tCtrl+S",
                            command=self._guardar_proyecto)
        m_arch.add_separator()
        m_arch.add_command(label="Exportar a Excel...\tCtrl+E",
                            command=self._exportar_excel)
        m_arch.add_separator()
        m_arch.add_command(label="Volver al menú",
                            command=self._volver_al_menu)

        # === CÁLCULO ===
        m_calc = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Cálculo", menu=m_calc)
        m_calc.add_command(label="Calcular Flujo DC\tF5",
                            command=self.calcular_flujo_dc)
        m_calc.add_separator()
        m_calc.add_command(label="Ver resultados",
                            command=self.mostrar_resultados)
        m_calc.add_command(label="Procedimiento paso a paso",
                            command=self.mostrar_procedimiento)

        # === AYUDA ===
        m_ayuda = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=m_ayuda)
        m_ayuda.add_command(label="Teoría del método DC",
                            command=self.mostrar_teoria)

    def __init__(self, root):
        self.root = root
        self.root.title("Vatia — Flujo DC Lineal")
        self.root.geometry("1500x900")
        self.root.minsize(1180, 720)
        self.root.configure(bg=self.C['bg'])

        # Listas de widgets
        self.tipo_widgets = []
        self.Pd_widgets = []
        self.Qd_widgets = []
        self.Pg_widgets = []
        self.Qg_widgets = []
        self.entries_lineas = []
        self.tipo_labels = []
        self._updating = False
        self._last_tipo_values = {}
        self.sidebar_labels = {}

        # Estado
        self.resultados = None
        self.BASE_MVA = 100.0
        self.V_BASE_KV = 220.0

                # Construir UI
        self._build_menu()
        self._build_header()
        self._build_toolbar()
        self._build_base_bar()
        self._build_main_area()
        self._build_action_bar()
        self._build_status_bar()

        # Atajos de teclado
        self.root.bind('<Control-n>', lambda e: self._nuevo_proyecto())
        self.root.bind('<Control-o>', lambda e: self._abrir_proyecto())
        self.root.bind('<Control-s>', lambda e: self._guardar_proyecto())
        self.root.bind('<Control-e>', lambda e: self._exportar_excel())
        self.root.bind('<F5>', lambda e: self.calcular_flujo_dc())

        # Cierre limpio
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.generar_tablas()

    # =================================================================
    # HEADER
    # =================================================================
    def _build_header(self):
        C = self.C
        hdr = tk.Frame(self.root, bg=C['primary_d'], height=90)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['amber'], height=4).pack(fill=tk.X, side=tk.TOP)

        left = tk.Frame(hdr, bg=C['primary_d'])
        left.pack(side=tk.LEFT, padx=24, pady=14)

        cv = tk.Canvas(left, width=54, height=54, bg=C['primary_d'],
                        highlightthickness=0)
        cv.pack(side=tk.LEFT, padx=(0, 14))
        cv.create_oval(2, 2, 52, 52, fill=C['amber'], outline='')
        cv.create_text(27, 27, text="DC", font=font(16, "bold"), fill='white')

        tinfo = tk.Frame(left, bg=C['primary_d'])
        tinfo.pack(side=tk.LEFT)
        tk.Label(tinfo, text="FLUJO DE POTENCIA LINEAL (DC)",
                 font=font(15, "bold"), bg=C['primary_d'], fg='white'
                 ).pack(anchor='w')
        tk.Label(tinfo,
                 text="Aproximación lineal  ·  Redes inductivas  ·  Sin iteraciones",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(anchor='w')

        right = tk.Frame(hdr, bg=C['primary_d'])
        right.pack(side=tk.RIGHT, padx=24, pady=15)
        badge = tk.Frame(right, bg=C['amber'])
        badge.pack()
        tk.Label(badge, text="  MÉTODO DC LINEAL  ",
                 font=font(10, "bold"), bg=C['amber'], fg='white',
                 padx=10, pady=7).pack()

    # =================================================================
    # TOOLBAR
    # =================================================================
    def _build_toolbar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(10, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['border'])
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel'])
        inner.pack(padx=18, pady=12, fill=tk.X)

        tk.Label(inner, text="Número de Nodos (N):", font=font(10, "bold"),
                 bg=C['panel'], fg=C['text']).pack(side=tk.LEFT)
        self.entry_n = ttk.Entry(inner, width=6, justify='center',
                                  font=font(11))
        self.entry_n.pack(side=tk.LEFT, padx=8, ipady=3)
        self.entry_n.insert(0, "3")

        tk.Button(inner, text="  Generar Tablas  ",
                  command=self.generar_tablas, bg=C['primary'], fg='white',
                  font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                  padx=16, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=10)
                # Botones de archivo
        tk.Button(inner, text="  Nuevo  ",
                  command=self._nuevo_proyecto,
                  bg=C['surface'], fg=C['text'],
                  font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                  padx=14, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=2)

        tk.Button(inner, text="  Abrir  ",
                  command=self._abrir_proyecto,
                  bg=C['surface'], fg=C['text'],
                  font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                  padx=14, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=2)

        tk.Button(inner, text="  Guardar  ",
                  command=self._guardar_proyecto,
                  bg=C['surface'], fg=C['text'],
                  font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                  padx=14, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=2)

        # Botón volver al menú
        tk.Button(inner, text="◄  Volver al menú",
                  command=self._volver_al_menu,
                  bg=C['surface'], fg=C['primary'],
                  font=font(10, "bold"), relief=tk.FLAT,
                  cursor='hand2', padx=16, pady=7, borderwidth=0
                  ).pack(side=tk.RIGHT, padx=2)

    # =================================================================
    # BASE BAR
    # =================================================================
    def _build_base_bar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(0, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['border'])
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel'])
        inner.pack(padx=18, pady=10, fill=tk.X)

        tk.Label(inner, text="DATOS DEL SISTEMA", font=font(9, "bold"),
                 bg=C['panel'], fg=C['primary']).pack(side=tk.LEFT, padx=(0, 20))

        tk.Label(inner, text="Potencia base:", font=font(9),
                 bg=C['panel'], fg=C['text_l']).pack(side=tk.LEFT)
        self.entry_base_mva = ttk.Entry(inner, width=8, justify='center',
                                          font=font(10))
        self.entry_base_mva.pack(side=tk.LEFT, padx=(6, 4), ipady=2)
        self.entry_base_mva.insert(0, "100")
        tk.Label(inner, text="MVA", font=font(9, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=(0, 24))

        tk.Label(inner, text="Voltaje base:", font=font(9),
                 bg=C['panel'], fg=C['text_l']).pack(side=tk.LEFT)
        self.entry_base_kv = ttk.Entry(inner, width=8, justify='center',
                                         font=font(10))
        self.entry_base_kv.pack(side=tk.LEFT, padx=(6, 4), ipady=2)
        self.entry_base_kv.insert(0, "220")
        tk.Label(inner, text="kV", font=font(9, "bold"),
                 bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=(0, 24))

        tk.Label(inner, text="→  Modelo: P_km = (θ_k − θ_m) / x_km",
                 font=font_m(9, "italic"), bg=C['panel'], fg=C['text_ll']
                 ).pack(side=tk.LEFT)

        self.entry_base_mva.bind('<KeyRelease>', lambda e: self._actualizar_bases())
        self.entry_base_kv.bind('<KeyRelease>', lambda e: self._actualizar_bases())

    def _actualizar_bases(self):
        try:
            self.BASE_MVA = self._parse_float(self.entry_base_mva.get(), 100.0)
            if self.BASE_MVA <= 0: self.BASE_MVA = 100.0
        except Exception:
            pass
        try:
            self.V_BASE_KV = self._parse_float(self.entry_base_kv.get(), 220.0)
            if self.V_BASE_KV <= 0: self.V_BASE_KV = 220.0
        except Exception:
            pass

    # =================================================================
    # MAIN AREA (sidebar + notebook + preview)
    # =================================================================
    def _build_main_area(self):
        C = self.C
        main = tk.Frame(self.root, bg=C['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)

        self.paned = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # Sidebar izquierda
        self.sidebar = tk.Frame(self.paned, bg=C['sidebar'], width=210)
        self.paned.add(self.sidebar, weight=0)
        self._build_sidebar()

        # Notebook al centro
        center = tk.Frame(self.paned, bg=C['bg'])
        self.paned.add(center, weight=5)
        self.notebook = ttk.Notebook(center)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.notebook.bind('<<NotebookTabChanged>>',
                            lambda e: self.root.after(30, self._sincronizar_tipos))

        # Preview derecha
        right = tk.Frame(self.paned, bg=C['bg'], width=380)
        self.paned.add(right, weight=4)
        right_lf = tk.Frame(right, bg=C['panel'], highlightthickness=1,
                             highlightbackground=C['border'])
        right_lf.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        tk.Label(right_lf, text="  VISTA PREVIA UNIFILAR  ",
                 font=font(10, "bold"), bg=C['head_bg'], fg=C['accent'],
                 pady=8).pack(fill=tk.X)

        self.preview_canvas = tk.Canvas(right_lf, bg=C['panel'],
                                          highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.preview_canvas.bind('<Configure>',
                                  lambda e: self._dibujar_preview())

    # =================================================================
    # SIDEBAR
    # =================================================================
    def _build_sidebar(self):
        C = self.C
        sb = self.sidebar
        sb.pack_propagate(False)

        tk.Label(sb, text="  PANEL DE CONTROL  ",
                 font=font(10, "bold"), bg=C['amber'], fg='white',
                 pady=10).pack(fill=tk.X)
        tk.Frame(sb, bg=C['accent'], height=2).pack(fill=tk.X)

        def _stat(key, label, color=None):
            frame = tk.Frame(sb, bg=C['sidebar'])
            frame.pack(fill=tk.X, padx=14, pady=(12, 4))
            tk.Label(frame, text=label.upper(), font=font(8, "bold"),
                     bg=C['sidebar'], fg=C['text_l']).pack(anchor='w')
            lbl = tk.Label(frame, text="—", font=font(18, "bold"),
                           bg=C['sidebar'], fg=color or C['text'])
            lbl.pack(anchor='w')
            self.sidebar_labels[key] = lbl

        def _sep():
            tk.Frame(sb, bg=C['border'], height=1).pack(fill=tk.X, padx=14, pady=8)

        _stat('nodos', 'Nodos del sistema', C['accent'])
        _stat('lineas', 'Líneas de transmisión', C['info'])
        _sep()
        _stat('slack', 'Nodos Slack', C['slack'])
        _stat('pv', 'Nodos PV', C['pv'])
        _stat('pq', 'Nodos PQ', C['pq'])
        _sep()
        _stat('estado', 'Estado del cálculo', C['text_l'])

        pie = tk.Frame(sb, bg=C['sidebar'])
        pie.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        tk.Label(pie, text="por Kevin", font=font(9, italic=True),
                 bg=C['sidebar'], fg=C['text_l']).pack()
        tk.Label(pie, text="Vatia © 2026", font=font_m(8),
                 bg=C['sidebar'], fg=C['text_ll']).pack()

        self.root.after(400, self._update_sidebar)

    def _update_sidebar(self):
        if not self.sidebar_labels:
            try:
                if self.root.winfo_exists():
                    self.root.after(500, self._update_sidebar)
            except Exception:
                pass
            return
        try:
            n = int(self.entry_n.get())
            n_lineas = 0
            for row in self.entries_lineas:
                try:
                    a = int(row[0].get()) if row[0].get().strip() else 0
                    b = int(row[1].get()) if row[1].get().strip() else 0
                    X = self._parse_float(row[2].get(), 0.0)
                    if a > 0 and b > 0 and abs(X) > 1e-9:
                        n_lineas += 1
                except Exception:
                    pass
            self.sidebar_labels['nodos'].config(text=str(n))
            self.sidebar_labels['lineas'].config(text=str(n_lineas))

            slack = pv = pq = 0
            for cb in self.tipo_widgets:
                t = cb.get()
                if "Slack" in t: slack += 1
                elif "PV" in t: pv += 1
                else: pq += 1
            self.sidebar_labels['slack'].config(text=str(slack))
            self.sidebar_labels['pv'].config(text=str(pv))
            self.sidebar_labels['pq'].config(text=str(pq))

            if self.resultados:
                self.sidebar_labels['estado'].config(
                    text="CALCULADO", fg=self.C['success'])
            else:
                self.sidebar_labels['estado'].config(
                    text="SIN CALCULAR", fg=self.C['text_l'])
        except Exception:
            pass

        try:
            if self.root.winfo_exists():
                self.root.after(500, self._update_sidebar)
        except Exception:
            pass

    # =================================================================
    # ACTION BAR
    # =================================================================
    def _build_action_bar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(6, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['border'])
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel'])
        inner.pack(padx=14, pady=10, fill=tk.X)

        def _grupo(parent, titulo):
            g = tk.Frame(parent, bg=C['panel'])
            g.pack(side=tk.LEFT, padx=(0, 20))
            tk.Label(g, text=titulo.upper(), font=font(8, "bold"),
                     bg=C['panel'], fg=C['text_l']).pack(anchor='w', pady=(0, 4))
            fila = tk.Frame(g, bg=C['panel'])
            fila.pack()
            return fila

        g1 = _grupo(inner, "Cálculo")
        self.btn_calcular = tk.Button(g1, text="  ▶  Calcular Flujo DC  ",
                  command=self.calcular_flujo_dc,
                  bg=C['amber'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=20, pady=7,
                  borderwidth=0)
        self.btn_calcular.pack(side=tk.LEFT, padx=2)

        g2 = _grupo(inner, "Vista")
        tk.Button(g2, text="  Ver Resultados  ",
                  command=self.mostrar_resultados,
                  bg=C['accent'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=16, pady=7,
                  borderwidth=0).pack(side=tk.LEFT, padx=2)

        tk.Button(g2, text="  Procedimiento paso a paso  ",
                  command=self.mostrar_procedimiento,
                  bg=C['info'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=16, pady=7,
                  borderwidth=0).pack(side=tk.LEFT, padx=2)

        g3 = tk.Frame(inner, bg=C['panel'])
        g3.pack(side=tk.RIGHT)
        tk.Label(g3, text="ARCHIVO / SOPORTE", font=font(8, "bold"),
                 bg=C['panel'], fg=C['text_l']).pack(anchor='e', pady=(0, 4))
        fila3 = tk.Frame(g3, bg=C['panel'])
        fila3.pack()

        tk.Button(fila3, text="  Exportar Excel  ",
                  command=self._exportar_excel,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=16, pady=7,
                  borderwidth=0).pack(side=tk.LEFT, padx=2)

        tk.Button(fila3, text="  Teoría del método  ",
                  command=self.mostrar_teoria,
                  bg=C['surface'], fg=C['primary'], font=font(10, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=16, pady=7,
                  borderwidth=0).pack(side=tk.LEFT, padx=2)
    # =================================================================
    # GUARDAR / ABRIR / NUEVO / EXPORTAR
    # =================================================================
    def _nuevo_proyecto(self):
        if messagebox.askyesno("Nuevo proyecto",
            "¿Crear un nuevo proyecto?\n\nSe perderán los datos actuales."):
            self.generar_tablas()
            self._set_status("Nuevo proyecto creado.", 'info')

    def _guardar_proyecto(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".vatia_dc",
            filetypes=[("Proyecto Vatia DC", "*.vatia_dc"),
                        ("Todos", "*.*")])
        if not path: return
        try:
            n = int(self.entry_n.get())
            data = {
                'version': '1.0',
                'modo': 'dc',
                'n': n,
                'base_mva': self.entry_base_mva.get(),
                'base_kv': self.entry_base_kv.get(),
                'nodos': [],
                'lineas': [],
            }
            for i in range(n):
                data['nodos'].append({
                    'tipo': self.tipo_widgets[i].get(),
                    'Pd': self.Pd_widgets[i].get(),
                    'Qd': self.Qd_widgets[i].get(),
                    'Pg': self.Pg_widgets[i].get(),
                    'Qg': self.Qg_widgets[i].get(),
                })
            for row in self.entries_lineas:
                data['lineas'].append([e.get() for e in row])

            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            self._set_status(f"Guardado: {os.path.basename(path)}", 'ok')
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))

    def _abrir_proyecto(self):
        from tkinter import filedialog
        import json
        path = filedialog.askopenfilename(
            filetypes=[("Proyecto Vatia DC", "*.vatia_dc"),
                        ("Proyecto Vatia", "*.vatia"),
                        ("Todos", "*.*")])
        if not path: return
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)

            # Bases
            if 'base_mva' in data:
                self.entry_base_mva.delete(0, tk.END)
                self.entry_base_mva.insert(0, data['base_mva'])
            if 'base_kv' in data:
                self.entry_base_kv.delete(0, tk.END)
                self.entry_base_kv.insert(0, data['base_kv'])
            self._actualizar_bases()

            # Nodos
            self.entry_n.delete(0, tk.END)
            self.entry_n.insert(0, str(data['n']))
            self.generar_tablas()

            for i, nd in enumerate(data.get('nodos', [])):
                if i >= len(self.tipo_widgets): break
                self.tipo_widgets[i].set(nd.get('tipo', 'PQ'))
                for w, key in [(self.Pd_widgets[i], 'Pd'),
                               (self.Qd_widgets[i], 'Qd'),
                               (self.Pg_widgets[i], 'Pg'),
                               (self.Qg_widgets[i], 'Qg')]:
                    w.delete(0, tk.END)
                    w.insert(0, nd.get(key, '0.0'))
                self.actualizar_fila_iter0(i)

            # Líneas
            for i, fila in enumerate(data.get('lineas', [])):
                if i >= len(self.entries_lineas): break
                for j, val in enumerate(fila):
                    if j >= len(self.entries_lineas[i]): break
                    self.entries_lineas[i][j].delete(0, tk.END)
                    self.entries_lineas[i][j].insert(0, val)

            # Refrescar todo
            self._last_tipo_values = {
                i: self.tipo_widgets[i].get()
                for i in range(min(data['n'], len(self.tipo_widgets)))
            }
            self._refresh_slack_options()
            for i in range(len(self.tipo_widgets)):
                self.actualizar_fila_iter0(i)
            self._dibujar_preview()
            self._update_sidebar()
            self._set_status(f"Cargado: {os.path.basename(path)}", 'ok')
        except Exception as e:
            messagebox.showerror("Error al abrir", str(e))

    def _exportar_excel(self):
        """Exporta datos y resultados del flujo DC a Excel."""
        if not self.resultados:
            if not messagebox.askyesno("Sin resultados",
                "Aún no se ha calculado el flujo DC.\n\n"
                "¿Exportar solo los datos de entrada?"):
                return
        try:
            self._set_status("Generando Excel...", 'info')
            self.root.update_idletasks()
            self._crear_excel_dc()
            self._set_status("Excel exportado correctamente.", 'ok')
        except Exception as e:
            import traceback
            messagebox.showerror("Error",
                f"Error al exportar:\n{e}\n\n{traceback.format_exc()}")

    def _crear_excel_dc(self):
        """Crea el archivo Excel con datos y resultados del flujo DC."""
        n = int(self.entry_n.get())
        BASE = self._parse_float(self.entry_base_mva.get(), 100.0)
        V_BASE = self._parse_float(self.entry_base_kv.get(), 220.0)

        wb = openpyxl.Workbook()

        # Estilos comunes
        font_titulo = Font(name="Calibri", size=14, bold=True,
                            color="FFFFFF")
        font_sec = Font(name="Calibri", size=12, bold=True,
                         color="FFFFFF")
        font_bold = Font(name="Calibri", size=10, bold=True)
        font_normal = Font(name="Calibri", size=10)
        font_mono = Font(name="Consolas", size=9)
        fill_titulo = PatternFill("solid", start_color="1E3A5F")
        fill_sec = PatternFill("solid", start_color="B45309")
        fill_head = PatternFill("solid", start_color="E2E8F0")
        fill_data = PatternFill("solid", start_color="F7F9FC")
        fill_calc = PatternFill("solid", start_color="FFF3E0")
        border_thin = Side(style="thin", color="000000")
        border_box = Border(left=border_thin, right=border_thin,
                            top=border_thin, bottom=border_thin)
        align_c = Alignment(horizontal="center", vertical="center")

        # ============ Hoja 1: Datos ============
        ws1 = wb.active
        ws1.title = "Datos DC"

        def sc1(r, c, val, font=None, fill=None, border=True, align=None):
            cell = ws1.cell(row=r, column=c, value=val)
            if font: cell.font = font
            if fill: cell.fill = fill
            if border: cell.border = border_box
            if align: cell.alignment = align
            return cell

        ws1.merge_cells("A1:H1")
        sc1(1, 1, "FLUJO DE POTENCIA LINEAL (DC) — DATOS Y RESULTADOS",
            font_titulo, fill_titulo, align=align_c)
        ws1.row_dimensions[1].height = 24

        r = 3
        ws1.merge_cells(f"A{r}:H{r}")
        sc1(r, 1, "1. BASES DEL SISTEMA", font_sec, fill_sec, align=align_c)
        r += 1
        sc1(r, 1, "S_base (MVA)", font_bold, fill_head, align=align_c)
        sc1(r, 2, BASE, font_normal, fill_data, align=align_c)
        sc1(r, 3, "V_base (kV)", font_bold, fill_head, align=align_c)
        sc1(r, 4, V_BASE, font_normal, fill_data, align=align_c)
        r += 2

        ws1.merge_cells(f"A{r}:H{r}")
        sc1(r, 1, "2. DATOS NODALES", font_sec, fill_sec, align=align_c)
        r += 1
        for c, h in enumerate(["Nodo", "Tipo", "Pd (MW)", "Qd (MVAr)",
                                "Pg (MW)", "Qg (MVAr)", "P_neta (MW)",
                                "P_neta (p.u.)"], 1):
            sc1(r, c, h, font_bold, fill_head, align=align_c)
        r += 1
        for i in range(n):
            tipo_str = self.tipo_widgets[i].get()
            Pd = self._parse_float(self.Pd_widgets[i].get(), 0.0)
            Qd = self._parse_float(self.Qd_widgets[i].get(), 0.0)
            Pg = self._parse_float(self.Pg_widgets[i].get(), 0.0)
            Qg = self._parse_float(self.Qg_widgets[i].get(), 0.0)
            P_neta = Pg - Pd
            fila = [i+1, tipo_str, Pd, Qd, Pg, Qg, P_neta, P_neta / BASE]
            for c, v in enumerate(fila, 1):
                sc1(r, c, v, font_normal, fill_data, align=align_c)
            r += 1
        r += 1

        ws1.merge_cells(f"A{r}:H{r}")
        sc1(r, 1, "3. DATOS DE LÍNEAS", font_sec, fill_sec, align=align_c)
        r += 1
        for c, h in enumerate(["Nodo A", "Nodo B", "X (p.u.)",
                                "B = 1/X (p.u.)"], 1):
            sc1(r, c, h, font_bold, fill_head, align=align_c)
        r += 1
        for row in self.entries_lineas:
            try:
                a = int(row[0].get()) if row[0].get().strip() else 0
                b = int(row[1].get()) if row[1].get().strip() else 0
                X = self._parse_float(row[2].get(), 0.0)
                if a > 0 and b > 0 and abs(X) > 1e-9:
                    for c, v in enumerate([a, b, X, 1.0/X], 1):
                        sc1(r, c, v, font_normal, fill_data, align=align_c)
                    r += 1
            except Exception:
                pass

        ws1.column_dimensions['A'].width = 14
        for c in "BCDEFGH":
            ws1.column_dimensions[c].width = 18

        # ============ Hoja 2: Resultados ============
        if self.resultados:
            ws2 = wb.create_sheet("Resultados DC")
            res = self.resultados
            types = res['tipos']

            def sc2(r, c, val, font=None, fill=None, border=True, align=None):
                cell = ws2.cell(row=r, column=c, value=val)
                if font: cell.font = font
                if fill: cell.fill = fill
                if border: cell.border = border_box
                if align: cell.alignment = align
                return cell

            ws2.merge_cells("A1:J1")
            sc2(1, 1, "RESULTADOS DEL FLUJO DE POTENCIA LINEAL (DC)",
                font_titulo, fill_titulo, align=align_c)
            ws2.row_dimensions[1].height = 24

            r2 = 3
            ws2.merge_cells(f"A{r2}:J{r2}")
            sc2(r2, 1, "ÁNGULOS NODALES", font_sec, fill_sec, align=align_c)
            r2 += 1
            for c, h in enumerate(["Nodo", "Tipo", "Pd (MW)", "Qd (MVAr)",
                                    "Pg (MW)", "Qg (MVAr)", "P_neta (MW)",
                                    "P_neta (p.u.)", "θ (rad)", "θ (°)"], 1):
                sc2(r2, c, h, font_bold, fill_head, align=align_c)
            r2 += 1
            for i in range(n):
                tipo_txt = {3: "Slack", 2: "PV", 1: "PQ"}[types[i]]
                P_iny = res['P_inj'][i]
                th = res['theta'][i]
                fila = [i+1, tipo_txt,
                        res['Pd'][i], res['Qd'][i],
                        res['Pg'][i], res['Qg'][i],
                        P_iny * BASE, P_iny,
                        th, math.degrees(th)]
                for c, v in enumerate(fila, 1):
                    sc2(r2, c, v, font_normal, fill_data, align=align_c)
                r2 += 1
            r2 += 1

            ws2.merge_cells(f"A{r2}:J{r2}")
            sc2(r2, 1, "FLUJOS POR LÍNEA", font_sec, fill_sec, align=align_c)
            r2 += 1
            for c, h in enumerate(["Línea", "Nodo A", "Nodo B", "X (p.u.)",
                                    "θ_A (°)", "θ_B (°)", "P_AB (p.u.)",
                                    "P_AB (MW)", "Sentido"], 1):
                sc2(r2, c, h, font_bold, fill_head, align=align_c)
            r2 += 1
            for k, F in enumerate(res['flujos']):
                P_ab = F['P']
                sentido = (f"{F['a']+1} → {F['b']+1}" if P_ab >= 0
                            else f"{F['b']+1} → {F['a']+1}")
                fila = [k+1, F['a']+1, F['b']+1, F['X'],
                        math.degrees(res['theta'][F['a']]),
                        math.degrees(res['theta'][F['b']]),
                        P_ab, P_ab * BASE, sentido]
                for c, v in enumerate(fila, 1):
                    sc2(r2, c, v, font_normal, fill_data, align=align_c)
                r2 += 1

            ws2.column_dimensions['A'].width = 12
            for c in "BCDEFGHIJ":
                ws2.column_dimensions[c].width = 15

        # Guardar
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        ruta = os.path.join(desktop, f"Flujo_DC_{n}_Nodos.xlsx")
        wb.save(ruta)
        try:
            os.startfile(ruta)
        except Exception:
            pass

        messagebox.showinfo("Éxito",
            f"Excel generado en:\n{ruta}\n\n"
            f"• Hoja 1: Datos de entrada\n"
            f"• Hoja 2: Resultados del flujo DC")

    def _on_close(self):
        try:
            self.root.destroy()
        except Exception:
            pass

    # =================================================================
    # STATUS BAR
    # =================================================================
    def _build_status_bar(self):
        C = self.C
        sb = tk.Frame(self.root, bg=C['primary_d'], height=32)
        sb.pack(fill=tk.X, side=tk.BOTTOM)
        sb.pack_propagate(False)
        tk.Frame(sb, bg=C['amber'], width=3).pack(side=tk.LEFT, fill=tk.Y)
        self.status_icon = tk.Label(sb, text="●", font=font(11),
                                      bg=C['primary_d'], fg=C['success'])
        self.status_icon.pack(side=tk.LEFT, padx=(12, 4))
        self.status_var = tk.StringVar(value="Listo.")
        tk.Label(sb, textvariable=self.status_var, font=font(9),
                 bg=C['primary_d'], fg='white').pack(side=tk.LEFT)
        tk.Label(sb, text="v1.0  ·  Módulo DC", font=font_m(8),
                 bg=C['primary_d'], fg=C['text_ll']).pack(side=tk.RIGHT, padx=(0, 14))

    def _set_status(self, text, kind='info'):
        try:
            self.status_var.set(text)
            colors = {'info': self.C['accent'], 'ok': self.C['success'],
                      'warn': self.C['warning'], 'err': self.C['danger']}
            self.status_icon.config(fg=colors.get(kind, self.C['accent']))
        except Exception:
            pass

    # =================================================================
    # GENERAR TABLAS
    # =================================================================
    def generar_tablas(self):
        try:
            n = int(self.entry_n.get())
            if n < 2:
                messagebox.showwarning("Advertencia", "Mínimo 2 nodos."); return
            if n > 20:
                messagebox.showwarning("Advertencia", "Máximo 20 nodos."); return
        except ValueError:
            messagebox.showwarning("Advertencia", "Ingrese un número entero."); return

        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)

        self.tipo_widgets.clear()
        self.Pd_widgets.clear()
        self.Qd_widgets.clear()
        self.Pg_widgets.clear()
        self.Qg_widgets.clear()
        self.entries_lineas.clear()
        self.tipo_labels.clear()
        self.resultados = None

        self.tab_nodos = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_lineas = tk.Frame(self.notebook, bg=self.C['panel'])
        self.notebook.add(self.tab_nodos, text='   Nodos (Pg, Pd)   ')
        self.notebook.add(self.tab_lineas, text='   Líneas (x)   ')

        self._crear_tab_nodos(n)
        self._crear_tab_lineas(n)

        # Nodo 1 = Slack por defecto
        self.tipo_widgets[0].set("Slack (ref.)")
        for i in range(1, n):
            self.tipo_widgets[i].set("PQ")
        self._last_tipo_values = {i: self.tipo_widgets[i].get()
                                   for i in range(n)}
        self._refresh_slack_options()
        for i in range(n):
            self.actualizar_fila_iter0(i)

        self._set_status(f"Sistema de {n} nodos configurado.", 'info')
        self._update_sidebar()
        self.root.after(120, self._dibujar_preview)

    def _crear_tab_nodos(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_nodos)
        

        hdr = tk.Frame(tab, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS NODALES — FLUJO DC",
                 font=font(12, "bold"), bg=C['primary_d'], fg='white',
                 pady=12).pack()
        tk.Label(hdr,
                 text="Tipo de nodo  ·  Potencias demandadas y generadas  ·  "
                      "Solo P se usa en el cálculo DC",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(tab, bg=C['panel'])
        grid.pack(padx=20, pady=20, fill=tk.X)

        headers = ["Nodo", "Tipo",
                   "Pd (MW)", "Qd (MVAr)",
                   "Pg (MW)", "Qg (MVAr)",
                   "P_neta (MW)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"),
                     bg=C['head_bg'], fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)

        for i in range(n):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']

            # Nodo
            tk.Label(grid, text=str(i+1), font=font(10, "bold"),
                     bg=bg, fg=C['accent'], padx=8, pady=6
                     ).grid(row=i+1, column=0, padx=1, pady=1, sticky='nsew')

            # Tipo
            cb = ttk.Combobox(grid, values=["Slack (ref.)", "PV", "PQ"],
                              state="readonly", font=font(9))
            cb.grid(row=i+1, column=1, padx=1, pady=1, sticky='nsew')
            cb.set("PQ")
            cb.bind("<<ComboboxSelected>>",
                    lambda e, idx=i: self._on_tipo_changed(idx))
            self.tipo_widgets.append(cb)

            # Pd
            e_pd = ttk.Entry(grid, justify='center', font=font(9))
            e_pd.grid(row=i+1, column=2, padx=1, pady=1, ipady=4, sticky='nsew')
            e_pd.insert(0, "0.0")
            e_pd.bind('<KeyRelease>',
                       lambda e, idx=i: self.actualizar_fila_iter0(idx))
            self.Pd_widgets.append(e_pd)

            # Qd
            e_qd = ttk.Entry(grid, justify='center', font=font(9))
            e_qd.grid(row=i+1, column=3, padx=1, pady=1, ipady=4, sticky='nsew')
            e_qd.insert(0, "0.0")
            self.Qd_widgets.append(e_qd)

            # Pg
            e_pg = ttk.Entry(grid, justify='center', font=font(9))
            e_pg.grid(row=i+1, column=4, padx=1, pady=1, ipady=4, sticky='nsew')
            e_pg.insert(0, "0.0")
            e_pg.bind('<KeyRelease>',
                       lambda e, idx=i: self.actualizar_fila_iter0(idx))
            self.Pg_widgets.append(e_pg)

            # Qg
            e_qg = ttk.Entry(grid, justify='center', font=font(9))
            e_qg.grid(row=i+1, column=5, padx=1, pady=1, ipady=4, sticky='nsew')
            e_qg.insert(0, "0.0")
            self.Qg_widgets.append(e_qg)

            # P_neta (etiqueta dinámica)
            lbl = tk.Label(grid, text="0.00", font=font_m(9),
                            bg=bg, fg=C['text_l'])
            lbl.grid(row=i+1, column=6, padx=1, pady=1, sticky='nsew')
            self.tipo_labels.append(lbl)

        # Nota informativa
        note = tk.Frame(tab, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        note.pack(fill=tk.X, padx=20, pady=(0, 20))
        tk.Label(note,
                 text="Nota:  El método DC usa únicamente la potencia activa "
                      "neta (P_neta = Pg − Pd).  Los valores de Qd y Qg se "
                      "almacenan pero no intervienen en el cálculo DC.",
                 font=font_m(9), bg=C['surface'], fg=C['text_l'],
                 justify='left', anchor='w', wraplength=1000
                 ).pack(fill=tk.X, padx=14, pady=10)
    def _crear_tab_lineas(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_lineas)

        hdr = tk.Frame(tab, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS DE LÍNEAS — FLUJO DC",
                 font=font(12, "bold"), bg=C['primary_d'], fg='white',
                 pady=12).pack()
        tk.Label(hdr, text="Solo se usa la reactancia X (p.u.)",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(tab, bg=C['panel'])
        grid.pack(padx=20, pady=20, fill=tk.X)

        headers = ["Nodo A", "Nodo B", "X (p.u.)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"),
                     bg=C['head_bg'], fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)

        for i in range(n):
            row = []
            for j in range(3):
                e = ttk.Entry(grid, justify='center', font=font(9))
                e.grid(row=i+1, column=j, padx=1, pady=1,
                       ipady=4, sticky='nsew')
                e.bind('<KeyRelease>', lambda ev: (
                    self._dibujar_preview(), self._update_sidebar()))
                e.bind('<FocusOut>', lambda ev: (
                    self._dibujar_preview(), self._update_sidebar()))
                row.append(e)
            self.entries_lineas.append(row)

    # =================================================================
    # SINCRONIZACIÓN DE TIPOS
    # =================================================================
    def _on_tipo_changed(self, idx):
        if self._updating: return
        nuevo_valor = self.tipo_widgets[idx].get()
        if self._last_tipo_values.get(idx) == nuevo_valor: return
        self._last_tipo_values[idx] = nuevo_valor
        self._updating = True
        try:
            self._refresh_slack_options()
            for i in range(len(self.tipo_widgets)):
                self.actualizar_fila_iter0(i)
            self._dibujar_preview()
            self._update_sidebar()
            self._set_status(f"Tipo del nodo {idx + 1} actualizado a {nuevo_valor}.")
        finally:
            self._updating = False

    def _refresh_slack_options(self):
        """Solo puede haber un nodo Slack en el sistema."""
        slack_idx = None
        for i, cb in enumerate(self.tipo_widgets):
            if "Slack" in cb.get():
                slack_idx = i
                break
        for i, cb in enumerate(self.tipo_widgets):
            current_val = cb.get()
            if i == slack_idx:
                new_vals = ["Slack (ref.)", "PV", "PQ"]
            elif slack_idx is not None:
                new_vals = ["PV", "PQ"]
            else:
                new_vals = ["Slack (ref.)", "PV", "PQ"]
            old_vals = list(cb['values']) if cb['values'] else []
            if old_vals != new_vals:
                cb.config(values=new_vals)
                if current_val in new_vals:
                    cb.set(current_val)
                elif current_val == "Slack (ref.)" and slack_idx is not None and i != slack_idx:
                    cb.set("PQ")
                    self._last_tipo_values[i] = "PQ"

    def _sincronizar_tipos(self):
        try:
            for i in range(len(self.tipo_widgets)):
                self.actualizar_fila_iter0(i)
        except Exception:
            pass

    def actualizar_fila_iter0(self, idx):
        """Actualiza la etiqueta de P_neta según Pg y Pd actuales."""
        if idx >= len(self.tipo_labels):
            return
        C = self.C
        try:
            Pg = self._parse_float(self.Pg_widgets[idx].get(), 0.0)
            Pd = self._parse_float(self.Pd_widgets[idx].get(), 0.0)
            P_neta = Pg - Pd
            lbl = self.tipo_labels[idx]
            lbl.config(text=f"{P_neta:+.2f}")
            if P_neta > 0.01:
                lbl.config(fg=C['success'])
            elif P_neta < -0.01:
                lbl.config(fg=C['pq'])
            else:
                lbl.config(fg=C['text_l'])
        except Exception:
            pass

    # =================================================================
    # PREVIEW
    # =================================================================
    def _dibujar_preview(self):
        if self._updating: return
        canvas = getattr(self, 'preview_canvas', None)
        if canvas is None: return
        try:
            n = int(self.entry_n.get())
        except Exception:
            return
        if n < 2 or len(self.tipo_widgets) != n: return

        canvas.delete("all")
        W = canvas.winfo_width(); H = canvas.winfo_height()
        if W < 50 or H < 50: W, H = 520, 520
        C = self.C

        canvas.create_rectangle(0, 0, W, H, fill='#FFFFFF', outline='')

        canvas.create_text(W / 2, 26, text="Diagrama Unifilar — Vista Previa",
                           font=font(13, "bold"), fill=C['primary'])
        canvas.create_line(W / 2 - 100, 42, W / 2 + 100, 42,
                           fill=C['amber'], width=2)

        tipos = []
        for cb in self.tipo_widgets:
            s = cb.get()
            tipos.append(3 if "Slack" in s else (2 if "PV" in s else 1))

        lineas = []
        for row in self.entries_lineas:
            a = self._parse_int(row[0].get(), 0)
            b = self._parse_int(row[1].get(), 0)
            X = self._parse_float(row[2].get(), 0.0)
            if a > 0 and b > 0 and a <= n and b <= n and a != b and abs(X) > 1e-9:
                lineas.append((a - 1, b - 1, X))

        if not lineas:
            canvas.create_text(
                W / 2, H / 2,
                text="Sin líneas definidas\n\n"
                     "Complete la pestaña 'Líneas (x)'\n"
                     "para visualizar el diagrama unifilar",
                font=font(11), fill=C['text_ll'], justify=tk.CENTER)
            return

        cx, cy = W / 2, H / 2 + 20
        R = min(W, H) * 0.30
        pos = []
        for i in range(n):
            ang = 2 * math.pi * i / n - math.pi / 2
            pos.append((cx + R * math.cos(ang), cy + R * math.sin(ang)))

        # Líneas
        for idx_l, (a, b, X) in enumerate(lineas):
            x1, y1 = pos[a]; x2, y2 = pos[b]
            canvas.create_line(x1 + 2, y1 + 2, x2 + 2, y2 + 2, width=4,
                               fill='#D5DCE3', capstyle=tk.ROUND)
            canvas.create_line(x1, y1, x2, y2, width=3,
                               fill=C['gray_dark'], capstyle=tk.ROUND)

            mx, my = (x1 + x2) / 2, (y1 + y2) / 2
            L = math.hypot(x2 - x1, y2 - y1) or 1.0
            perp_x = -(y2 - y1) / L
            perp_y = (x2 - x1) / L
            offset = 22 if idx_l % 2 == 0 else -22
            lx_lbl = mx + perp_x * offset
            ly_lbl = my + perp_y * offset

            lbl = f"x={X:.4f}"
            tid = canvas.create_text(lx_lbl, ly_lbl, text=lbl,
                                     font=font_m(8, "bold"),
                                     fill=C['amber'], justify=tk.CENTER)
            bb = canvas.bbox(tid)
            if bb:
                canvas.create_rectangle(bb[0] - 4, bb[1] - 2,
                                         bb[2] + 4, bb[3] + 2,
                                         fill='#FFFFFF', outline=C['amber'],
                                         width=1)
                canvas.tag_raise(tid)
            canvas.create_line(mx, my, lx_lbl, ly_lbl,
                               fill=C['border'], width=1, dash=(2, 2))

        # Nodos
        color_map = {1: C['pq'], 2: C['pv'], 3: C['slack']}
        ring_map = {1: '#FFEDD5', 2: '#EDE9FE', 3: '#D1FAE5'}
        name_map = {1: 'PQ', 2: 'PV', 3: 'SLK'}
        for i, (x, y) in enumerate(pos):
            tipo = tipos[i]
            col = color_map[tipo]; ring = ring_map[tipo]
            canvas.create_line(x - 24, y, x + 24, y, width=10,
                               fill=C['gray_dark'], capstyle=tk.ROUND)
            canvas.create_oval(x - 17, y - 17, x + 17, y + 17, fill=ring, outline='')
            canvas.create_oval(x - 13, y - 13, x + 13, y + 13, fill=col,
                               outline='white', width=2)
            canvas.create_text(x, y, text=str(i + 1), fill='white',
                               font=font(10, "bold"))
            canvas.create_text(x, y + 30, text=name_map[tipo],
                               font=font(8, "bold"), fill=col)

        # Leyenda
        lx, ly = 12, H - 82
        canvas.create_rectangle(lx + 2, ly + 2, lx + 158, ly + 70,
                                fill='#D5DCE3', outline='')
        canvas.create_rectangle(lx, ly, lx + 156, ly + 68, fill=C['panel'],
                                outline=C['border'], width=1)
        canvas.create_text(lx + 78, ly + 14, text="LEYENDA",
                           font=font(8, "bold"), fill=C['amber'])
        canvas.create_line(lx + 40, ly + 24, lx + 116, ly + 24, fill=C['border'])
        items = [(C['slack'], 'Slack'), (C['pv'], 'PV'), (C['pq'], 'PQ')]
        for k, (col, txt) in enumerate(items):
            yk = ly + 32 + k * 12
            canvas.create_oval(lx + 14, yk - 5, lx + 24, yk + 5, fill=col,
                               outline='white', width=1)
            canvas.create_text(lx + 32, yk, text=txt, font=font(8),
                               fill=C['text'], anchor='w')

        canvas.create_text(W - 14, H - 12,
                           text=f"{n} nodos · {len(lineas)} línea(s)",
                           font=font(8, italic=True), fill=C['text_l'],
                           anchor='se')

    # =================================================================
    # CÁLCULO DEL FLUJO DC
    # =================================================================
    def calcular_flujo_dc(self):
        try:
            n = int(self.entry_n.get())
            self._actualizar_bases()

            tipos = []; Pg = []; Pd = []; Qg = []; Qd = []
            for i in range(n):
                t = self.tipo_widgets[i].get()
                tipos.append(3 if "Slack" in t else (2 if "PV" in t else 1))
                Pd.append(self._parse_float(self.Pd_widgets[i].get(), 0.0))
                Qd.append(self._parse_float(self.Qd_widgets[i].get(), 0.0))
                Pg.append(self._parse_float(self.Pg_widgets[i].get(), 0.0))
                Qg.append(self._parse_float(self.Qg_widgets[i].get(), 0.0))

            # Validar único Slack
            n_slack = sum(1 for t in tipos if t == 3)
            if n_slack != 1:
                messagebox.showerror("Error",
                    f"Debe haber exactamente 1 nodo Slack. Hay {n_slack}.")
                return

            lineas = []
            for row in self.entries_lineas:
                try:
                    a = int(row[0].get()) if row[0].get().strip() else 0
                    b = int(row[1].get()) if row[1].get().strip() else 0
                    X = self._parse_float(row[2].get(), 0.0)
                    if (a > 0 and b > 0 and a <= n and b <= n
                            and a != b and abs(X) > 1e-9):
                        lineas.append({'a': a-1, 'b': b-1, 'X': X})
                except Exception:
                    continue

            if not lineas:
                messagebox.showerror("Error", "No hay líneas válidas definidas.")
                return

            # Vector de inyecciones netas P (en p.u.)
            P_inj = [(Pg[i] - Pd[i]) / self.BASE_MVA for i in range(n)]

            # Índices de nodos no-slack (PV y PQ son equivalentes en DC)
            no_slack = [i for i in range(n) if tipos[i] != 3]
            idx_map = {orig: k for k, orig in enumerate(no_slack)}
            m = len(no_slack)
            if m == 0:
                messagebox.showerror("Error", "No hay nodos PV/PQ definidos.")
                return

            # Construir matriz B (m x m)
            B = [[0.0]*m for _ in range(m)]
            for L in lineas:
                a, b, X = L['a'], L['b'], L['X']
                y = 1.0 / X
                if a in idx_map: B[idx_map[a]][idx_map[a]] += y
                if b in idx_map: B[idx_map[b]][idx_map[b]] += y
                if a in idx_map and b in idx_map:
                    B[idx_map[a]][idx_map[b]] -= y
                    B[idx_map[b]][idx_map[a]] -= y

            # Vector P para nodos no-slack
            P_vec = [P_inj[i] for i in no_slack]

            # Resolver B · θ = P
            B_np = np.array(B, dtype=np.float64)
            P_np = np.array(P_vec, dtype=np.float64)
            try:
                theta_no_slack = np.linalg.solve(B_np, P_np)
            except np.linalg.LinAlgError as e:
                messagebox.showerror("Error", f"Matriz B singular: {e}")
                return

            # Ángulos completos (Slack = 0)
            theta = np.zeros(n)
            for i, orig in enumerate(no_slack):
                theta[orig] = theta_no_slack[i]

            # Flujos por línea
            flujos = []
            for L in lineas:
                a, b, X = L['a'], L['b'], L['X']
                P_ab = (theta[a] - theta[b]) / X
                flujos.append({'a': a, 'b': b, 'P': P_ab, 'X': X})

            self.resultados = {
                'n': n, 'tipos': tipos,
                'Pg': Pg, 'Pd': Pd, 'P_inj': P_inj,
                'Qg': Qg, 'Qd': Qd,
                'theta': theta, 'lineas': lineas, 'flujos': flujos,
                'B': B, 'no_slack': no_slack,
            }
            self._set_status("Flujo DC calculado correctamente.", 'ok')
            self._update_sidebar()
            self.mostrar_resultados()
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Error en el cálculo:\n{e}")

    @staticmethod
    def _parse_float(s, default=0.0):
        try:
            if s is None: return default
            s = str(s).strip().replace(',', '.')
            return float(s) if s else default
        except Exception:
            return default

    @staticmethod
    def nodo_label(idx):
        """Devuelve la etiqueta del nodo con subíndice Unicode: _1 → ₁"""
        subs = '₀₁₂₃₄₅₆₇₈₉'
        s = str(idx + 1)
        return ''.join(subs[int(d)] for d in s)

    @staticmethod
    def _parse_int(s, default=0):
        try:
            if s is None: return default
            s = str(s).strip().replace(',', '.')
            return int(float(s)) if s else default
        except Exception:
            return default

    def _crear_tab_scrolleable(self, parent):
        """Frame con scroll vertical dentro de una pestaña."""
        canvas = tk.Canvas(parent, bg=self.C['panel'], highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient='vertical',
                                  command=canvas.yview)
        interior = tk.Frame(canvas, bg=self.C['panel'])
        interior.bind('<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=interior, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<Configure>',
            lambda e: canvas.itemconfig(canvas.find_all()[0], width=e.width))

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        def _bind_wheel(e):
            canvas.bind_all('<MouseWheel>', _on_mousewheel)
        def _unbind_wheel(e):
            canvas.unbind_all('<MouseWheel>')
        canvas.bind('<Enter>', _bind_wheel)
        canvas.bind('<Leave>', _unbind_wheel)
        return interior
    
    def _matriz_widget(self, parent, valores, color=None, ancho_celda=110):
        """Devuelve un Canvas con una matriz dibujada con corchetes [ ]."""
        C = self.C
        if color is None:
            color = C['primary']
        n_rows = len(valores)
        n_cols = len(valores[0]) if valores else 0
        if n_rows == 0 or n_cols == 0:
            return tk.Frame(parent, bg=C['bg'])

        cell_h = 32
        bracket_w = 20
        W = bracket_w * 2 + n_cols * ancho_celda
        H = n_rows * cell_h

        cv = tk.Canvas(parent, width=W, height=H,
                       bg=C['bg'], highlightthickness=0)

        for i in range(n_rows):
            for j in range(n_cols):
                x0 = bracket_w + j * ancho_celda
                y0 = i * cell_h
                bg_cell = C['panel'] if i % 2 == 0 else C['surface']
                cv.create_rectangle(x0, y0, x0 + ancho_celda, y0 + cell_h,
                                     fill=bg_cell, outline='')
                cv.create_text(x0 + ancho_celda / 2, y0 + cell_h / 2,
                               text=str(valores[i][j]),
                               font=font_m(10), fill=C['text'])

        # Corchete izquierdo [
        cv.create_line(bracket_w - 6, 2, 6, 2, fill=color, width=2)
        cv.create_line(6, 2, 6, H - 2, fill=color, width=2)
        cv.create_line(6, H - 2, bracket_w - 6, H - 2, fill=color, width=2)
        # Corchete derecho ]
        xr = bracket_w + n_cols * ancho_celda
        cv.create_line(xr + 6, 2, xr + bracket_w - 6, 2, fill=color, width=2)
        cv.create_line(xr + bracket_w - 6, 2, xr + bracket_w - 6, H - 2,
                       fill=color, width=2)
        cv.create_line(xr + bracket_w - 6, H - 2, xr + 6, H - 2,
                       fill=color, width=2)

        return cv
    
    def mostrar_procedimiento(self):
        """Muestra el planteamiento paso a paso del método DC."""
        if not self.resultados:
            messagebox.showinfo("Info",
                "Calcula primero el flujo DC para ver el procedimiento.")
            return

        r = self.resultados
        n = r['n']
        tipos = r['tipos']
        Pg = r['Pg']; Pd = r['Pd']
        theta = r['theta']
        lineas = r['lineas']
        flujos = r['flujos']
        no_slack = r['no_slack']
        BASE = self.BASE_MVA

        # ============ Reconstruir matrices necesarias ============
        # Matriz B completa (n x n)
        B_full = [[0.0] * n for _ in range(n)]
        for L in lineas:
            a, b, X = L['a'], L['b'], L['X']
            y = 1.0 / X
            B_full[a][a] += y
            B_full[b][b] += y
            B_full[a][b] -= y
            B_full[b][a] -= y

        # Matriz B reducida (sin slack)
        m = len(no_slack)
        B_red = [[B_full[i][j] for j in no_slack] for i in no_slack]

        # Vector de inyecciones P (p.u.)
        P_vec = [(Pg[i] - Pd[i]) / BASE for i in no_slack]

        # Inversa
        try:
            B_inv = np.linalg.inv(np.array(B_red, dtype=np.float64))
        except Exception:
            B_inv = None

        # ============ Ventana ============
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Procedimiento — Flujo DC Lineal")
        win.geometry("1100x820")
        win.minsize(950, 700)
        win.configure(bg=C['bg'])

        hdr = tk.Frame(win, bg=C['primary_d'], height=70)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['amber'], height=4).pack(fill=tk.X, side=tk.TOP)
        tk.Label(hdr,
                 text="  PROCEDIMIENTO — FLUJO DE POTENCIA LINEAL (DC)",
                 font=font(14, "bold"), bg=C['primary_d'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=22)
        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['danger'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=14, pady=6, cursor='hand2',
                  borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=22)

        # Scrollable
        canvas = tk.Canvas(win, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(win, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas.find_all()[0],
                                                 width=e.width))
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        def _wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        canvas.bind('<MouseWheel>', _wheel)
        inner.bind('<MouseWheel>', _wheel)

        # ============ Helpers locales ============
        def seccion(numero, titulo):
            """Encabezado de sección numerada."""
            f = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                          highlightbackground=C['border'])
            f.pack(fill=tk.X, padx=30, pady=(24, 10))
            tk.Frame(f, bg=C['amber'], width=5).pack(side=tk.LEFT, fill=tk.Y)
            box = tk.Frame(f, bg=C['amber'], width=34, height=34)
            box.pack(side=tk.LEFT, padx=10, pady=8)
            box.pack_propagate(False)
            tk.Label(box, text=str(numero), font=font(14, "bold"),
                     bg=C['amber'], fg='white').pack(expand=True)
            tk.Label(f, text=titulo, font=font(12, "bold"),
                     bg=C['surface'], fg=C['text'],
                     padx=8, pady=10).pack(side=tk.LEFT)

        def parrafo(txt, mono=False):
            tk.Label(inner, text=txt,
                     font=font_m(10) if mono else font(10),
                     bg=C['bg'], fg=C['text'], justify='left',
                     anchor='w', wraplength=1000
                     ).pack(fill=tk.X, padx=34, pady=3)

        def formula(txt):
            """Caja de fórmula destacada."""
            f = tk.Frame(inner, bg=C['panel'], highlightthickness=1,
                          highlightbackground=C['border'])
            f.pack(fill=tk.X, padx=50, pady=6)
            tk.Label(f, text=txt, font=font_m(11), bg=C['panel'],
                     fg=C['primary'], justify='left', anchor='w',
                     padx=16, pady=10).pack(fill=tk.X)

        def matriz_frame(parent):
            """Frame donde se pegará una matriz."""
            f = tk.Frame(parent, bg=C['bg'])
            f.pack(anchor='w', padx=40, pady=6)
            return f

        # ========================================================
        # SECCIÓN 1 — SUSCEPTANCIAS
        # ========================================================
        seccion(1, "CÁLCULO DE SUSCEPTANCIAS  (B_km = 1 / x_km)")
        parrafo("Para cada línea se calcula su susceptancia serie "
                "como el inverso de su reactancia:")
        formula("B_km = 1 / x_km   [p.u.]")
        parrafo("Valores para cada línea del sistema:", mono=True)
        for k, L in enumerate(lineas):
            X = L['X']
            B_km = 1.0 / X
            formula(f"Línea {k+1}  (nodo {L['a']+1} – {L['b']+1}):   "
                    f"B = 1 / {X:.4f} = {B_km:.4f}  p.u.")

        # ========================================================
        # SECCIÓN 2 — MATRIZ B COMPLETA
        # ========================================================
        seccion(2, f"MATRIZ B COMPLETA  ({n} × {n})")
        parrafo("Se construye una matriz cuadrada de tamaño n × n con:")
        parrafo("• Diagonal:  B_kk = Σ 1/x_km  (suma de susceptancias "
                "conectadas al nodo k)")
        parrafo("• Fuera de diagonal:  B_km = −1/x_km  si existe línea; 0 si no")

        B_full_str = [[f"{B_full[i][j]:+.4f}" for j in range(n)]
                      for i in range(n)]
        mf1 = matriz_frame(inner)
        self._matriz_widget(mf1, B_full_str, color=C['primary']).pack()

        # Indicar cuál es el slack
        slack_idx = 0
        for i in range(n):
            if tipos[i] == 3:
                slack_idx = i
                break
        parrafo(f"El nodo Slack es el {slack_idx + 1}. "
                f"Se eliminará su fila y columna para construir B reducida.")

        # ========================================================
        # SECCIÓN 3 — MATRIZ B REDUCIDA (sin slack)
        # ========================================================
        seccion(3, f"MATRIZ B REDUCIDA  ({m} × {m})  —  sin el nodo Slack")
        parrafo("Al eliminar el nodo Slack, se obtiene una matriz cuadrada "
                "de tamaño (n-1) × (n-1) que relaciona las inyecciones netas "
                "de los demás nodos con sus ángulos.")

        B_red_str = [[f"{B_red[i][j]:+.4f}" for j in range(m)]
                     for i in range(m)]
        mf2 = matriz_frame(inner)
        self._matriz_widget(mf2, B_red_str, color=C['amber']).pack()

        # ========================================================
        # SECCIÓN 4 — VECTOR DE INYECCIONES
        # ========================================================
        seccion(4, "VECTOR DE INYECCIONES NETAS  ΔP")
        parrafo("Se calcula la inyección neta de cada nodo no-Slack como:")
        formula("P_neta = (Pg − Pd) / S_base   [p.u.]")
        parrafo("Donde Pg es la potencia generada y Pd la demandada, ambas "
                "en MW. S_base es la potencia base del sistema.")

        for k, orig in enumerate(no_slack):
            Pg_i = Pg[orig]; Pd_i = Pd[orig]
            P_pu = (Pg_i - Pd_i) / BASE
            formula(f"P_neta{self.nodo_label(orig)} = ({Pg_i:+.2f} − {Pd_i:+.2f}) / {BASE:.1f}"
                    f" = {P_pu:+.4f}  p.u.")

        # Mostrar vector como columna
        P_col = [[f"{p:+.4f}"] for p in P_vec]
        mf3 = matriz_frame(inner)
        self._matriz_widget(mf3, P_col, color=C['accent']).pack()

        # ========================================================
        # SECCIÓN 5 — ECUACIÓN MATRICIAL
        # ========================================================
        seccion(5, "ECUACIÓN MATRICIAL DEL SISTEMA DC")
        formula("B_red · θ = ΔP")
        parrafo("Se despeja el vector de ángulos multiplicando ambos lados "
                "por la inversa de B_red:")
        formula("θ = B_red⁻¹ · ΔP")

        # ========================================================
        # SECCIÓN 6 — INVERSA DE B
        # ========================================================
        seccion(6, f"INVERSA DE LA MATRIZ B  ({m} × {m})")
        if B_inv is not None:
            B_inv_str = [[f"{B_inv[i][j]:+.4f}" for j in range(m)]
                         for i in range(m)]
            mf4 = matriz_frame(inner)
            self._matriz_widget(mf4, B_inv_str, color=C['success']).pack()
        else:
            parrafo("⚠ No se pudo calcular la inversa (matriz singular).",
                    mono=True)

        # ========================================================
        # SECCIÓN 7 — SOLUCIÓN: ÁNGULOS NODALES
        # ========================================================
        seccion(7, "SOLUCIÓN — ÁNGULOS NODALES θ")
        parrafo("Se multiplica la inversa de B por el vector ΔP para "
                "obtener los ángulos en cada nodo no-Slack:", mono=False)

        for k, orig in enumerate(no_slack):
            th_rad = theta[orig]
            th_deg = math.degrees(th_rad)
            formula(f"θ{self.nodo_label(orig)} = {th_rad:+.6f} rad"
                    f"  =  {th_deg:+.4f}°")

        parrafo(f"El nodo Slack mantiene su ángulo fijo:  "
                f"θ{self.nodo_label(slack_idx)} = 0 rad  (referencia del sistema).")

        # Vector θ completo
        theta_col = [[f"{theta[i]:+.6f}"] for i in range(n)]
        mf5 = matriz_frame(inner)
        self._matriz_widget(mf5, theta_col, color=C['primary']).pack()

        # ========================================================
        # SECCIÓN 8 — FLUJOS POR LÍNEA
        # ========================================================
        seccion(8, "FLUJOS POR LÍNEA  —  P_L = (θ_k − θ_m) / x_km")
        parrafo("Con los ángulos obtenidos se calcula el flujo de potencia "
                "activa por cada línea:")

        for k, F in enumerate(flujos):
            th_a = theta[F['a']]
            th_b = theta[F['b']]
            X = F['X']
            P_ab = (th_a - th_b) / X
            sentido = (f"{F['a']+1} → {F['b']+1}" if P_ab >= 0
                        else f"{F['b']+1} → {F['a']+1}")
            formula(f"Línea {k+1}:  P = ({th_a:+.6f} − {th_b:+.6f}) / {X:.4f}  "
                    f"= {P_ab:+.4f} p.u.  =  {P_ab*BASE:+.2f} MW   ({sentido})")

        # ========================================================
        # SECCIÓN 9 — BALANCE DE POTENCIA
        # ========================================================
        seccion(9, "BALANCE DE POTENCIA DEL SISTEMA")
        parrafo("En el modelo DC no se consideran pérdidas, por lo que la "
                "generación total debe igualar a la demanda total:")
        formula("Σ Pg  =  Σ Pd  +  pérdidas     donde    pérdidas = 0  (DC)")

        P_gen_total = sum(Pg)
        P_load_total = sum(Pd)
        P_slack_iny = P_load_total - P_gen_total + Pg[slack_idx] * 0
         # En realidad: la inyección del slack es lo que balancea
         # P_slack + Σ Pg_otros = Σ Pd
        P_gen_otros = sum(Pg[i] for i in range(n) if i != slack_idx)
        P_slack_gen = P_load_total - P_gen_otros

        formula(f"Σ Pg,otros = {P_gen_otros:.2f} MW    ·    "
                f"Σ Pd = {P_load_total:.2f} MW")
        formula(f"Pg,Slack (necesario) = Σ Pd − Σ Pg,otros = "
                f"{P_slack_gen:.2f} MW")
        formula(f"Verificación:  Σ Pg = {P_gen_otros + P_slack_gen:.2f} MW "
                f"= Σ Pd = {P_load_total:.2f} MW   ✓")

        # ========================================================
        # Nota
        # ========================================================
        note = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        note.pack(fill=tk.X, padx=30, pady=(20, 24))
        tk.Label(note, text="OBSERVACIONES",
                 font=font(9, "bold"), bg=C['surface'],
                 fg=C['text'], anchor='w').pack(fill=tk.X, padx=14,
                                                  pady=(10, 2))
        tk.Label(note,
                 text="• El método DC ignora las pérdidas y las tensiones.\n"
                      "• El nodo Slack balancea el sistema: su generación "
                      "cubre lo que los demás nodos no alcanzan a generar.\n"
                      "• Los ángulos están referidos al Slack (θ_Slack = 0).\n"
                      "• Los flujos positivos indican potencia inyectada del "
                      "nodo A al B; negativos son el caso contrario.",
                 font=font_m(9), bg=C['surface'], fg=C['text_l'],
                 justify='left', anchor='w', wraplength=1000
                 ).pack(fill=tk.X, padx=14, pady=(0, 12))

        # Botón cerrar
        tk.Button(inner, text="Cerrar ventana", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=30, pady=10, cursor='hand2',
                  borderwidth=0).pack(pady=(0, 24))


    # =================================================================
    # MOSTRAR RESULTADOS
    # =================================================================
    def mostrar_resultados(self):
        if not self.resultados:
            messagebox.showinfo("Info", "Calcula primero el flujo DC.")
            return

        r = self.resultados
        n = r['n']; tipos = r['tipos']
        Pg = r['Pg']; Pd = r['Pd']; P_inj = r['P_inj']
        Qg = r.get('Qg', [0.0]*n); Qd = r.get('Qd', [0.0]*n)
        theta = r['theta']; flujos = r['flujos']
        BASE = self.BASE_MVA

        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Resultados — Flujo DC Lineal")
        win.geometry("1400x850")
        win.minsize(1200, 700)
        win.configure(bg=C['bg'])

        # === Header ===
        hdr = tk.Frame(win, bg=C['primary_d'], height=70)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['amber'], height=4).pack(fill=tk.X, side=tk.TOP)
        tk.Label(hdr, text="  RESULTADOS — FLUJO DE POTENCIA LINEAL (DC)",
                 font=font(15, "bold"), bg=C['primary_d'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=22)
        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['danger'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=14, pady=6, cursor='hand2',
                  borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=22)

        # === Scroll ===
        canvas = tk.Canvas(win, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(win, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas.find_all()[0],
                                                 width=e.width))
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        def _wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        canvas.bind('<MouseWheel>', _wheel)
        inner.bind('<MouseWheel>', _wheel)

        # === Info de bases ===
        info_base = tk.Frame(inner, bg=C['surface'],
                             highlightthickness=1,
                             highlightbackground=C['border'])
        info_base.pack(fill=tk.X, padx=30, pady=(20, 18))
        tk.Label(info_base,
                 text=f"Base de potencia: S_base = {BASE:.1f} MVA    ·    "
                      f"Slack con θ = 0    ·    "
                      f"Total {n} nodos, {len(flujos)} líneas",
                 font=font_m(10), bg=C['surface'], fg=C['text'],
                 padx=14, pady=10, anchor='w').pack(fill=tk.X)

        # ============================================================
        # SECCIÓN 1: ÁNGULOS NODALES
        # ============================================================
        self._titulo_seccion_resultados(inner, "ÁNGULOS NODALES", C['primary'])

        # Tabla con columnas fijas
        cols_ang = [
            ("Nodo",           8),
            ("Tipo",          10),
            ("Pd (MW)",       11),
            ("Qd (MVAr)",     11),
            ("Pg (MW)",       11),
            ("Qg (MVAr)",     11),
            ("P_neta (MW)",   13),
            ("P_neta (p.u.)", 14),
            ("θ (rad)",       14),
            ("θ (°)",         12),
        ]

        tabla = tk.Frame(inner, bg=C['border'])
        tabla.pack(fill=tk.X, padx=30, pady=(0, 22))

        # Header
        head = tk.Frame(tabla, bg=C['primary_d'])
        head.pack(fill=tk.X)
        for i, (h, w) in enumerate(cols_ang):
            tk.Label(head, text=h, font=font(9, "bold"),
                     bg=C['primary_d'], fg='white',
                     width=w, pady=8).pack(side=tk.LEFT, padx=0)

        # Filas
        body = tk.Frame(tabla, bg=C['border'])
        body.pack(fill=tk.X)
        for i in range(n):
            bg_row = C['row_even'] if i % 2 == 0 else C['row_odd']
            tipo_txt = {3: "Slack", 2: "PV", 1: "PQ"}[tipos[i]]

            valores = [
                (f"{i+1}",                     bg_row, C['text']),
                (tipo_txt,                     
                 {3: C['slack'], 2: C['pv'], 1: C['pq']}[tipos[i]],
                 'white'),
                (f"{Pd[i]:+.2f}",              bg_row, C['text']),
                (f"{Qd[i]:+.2f}",              bg_row, C['text']),
                (f"{Pg[i]:+.2f}",              bg_row, C['text']),
                (f"{Qg[i]:+.2f}",              bg_row, C['text']),
                (f"{P_inj[i]*BASE:+.2f}",      bg_row, C['accent']),
                (f"{P_inj[i]:+.4f}",           bg_row, C['text']),
                (f"{theta[i]:+.4f}",           bg_row, C['primary']),
                (f"{math.degrees(theta[i]):+.3f}", bg_row, C['primary']),
            ]
            fila = tk.Frame(body, bg=C['border'])
            fila.pack(fill=tk.X, pady=1)
            for (v, w), (_, col_bg, col_fg) in zip(cols_ang, valores):
                tk.Label(fila, text=v, font=font_m(9),
                         bg=col_bg, fg=col_fg,
                         width=w, pady=6).pack(side=tk.LEFT, padx=0)

        # ============================================================
        # SECCIÓN 2: RESUMEN DE POTENCIA
        # ============================================================
        self._titulo_seccion_resultados(inner,
            "BALANCE DE POTENCIA", C['primary'])

        P_gen_total = sum(Pg)
        P_load_total = sum(Pd)
        Q_gen_total = sum(Qg)
        Q_load_total = sum(Qd)
        # En DC no se calculan pérdidas reales, solo inyecciones netas
        P_inj_total = sum(P_inj) * BASE
        # En DC no hay pérdidas: la suma de inyecciones de nodos NO-slack
        # es igual (con signo opuesto) a la inyección del nodo Slack.
        P_slack_iny = -P_inj_total

        resumen = tk.Frame(inner, bg=C['bg'])
        resumen.pack(fill=tk.X, padx=30, pady=(0, 22))

        def _card(parent, titulo, valor, sub, color):
            card = tk.Frame(parent, bg=C['panel'], highlightthickness=1,
                            highlightbackground=C['border'])
            card.pack(side=tk.LEFT, padx=8, fill=tk.Y)
            tk.Frame(card, bg=color, width=5).pack(side=tk.LEFT, fill=tk.Y)
            body = tk.Frame(card, bg=C['panel'])
            body.pack(side=tk.LEFT, padx=16, pady=12)
            tk.Label(body, text=titulo.upper(),
                     font=font(8, "bold"), bg=C['panel'],
                     fg=C['text_l']).pack(anchor='w')
            tk.Label(body, text=valor, font=font(14, "bold"),
                     bg=C['panel'], fg=color).pack(anchor='w', pady=(4, 0))
            tk.Label(body, text=sub, font=font(8, italic=True),
                     bg=C['panel'], fg=C['text_l']).pack(anchor='w')

        _card(resumen, "Generación total P",
              f"{P_gen_total:+.2f} MW",
              f"Q: {Q_gen_total:+.2f} MVAr", C['slack'])
        _card(resumen, "Demanda total P",
              f"{P_load_total:+.2f} MW",
              f"Q: {Q_load_total:+.2f} MVAr", C['pq'])
        _card(resumen, "Inyección del Slack (calculada)",
              f"{P_slack_iny:+.2f} MW",
              "Balance del sistema (sin pérdidas en DC)", C['amber'])

        # ============================================================
        # SECCIÓN 3: FLUJOS POR LÍNEA
        # ============================================================
        self._titulo_seccion_resultados(inner,
            "FLUJOS POR LÍNEA", C['primary'])

        cols_fl = [
            ("Línea",        9),
            ("Nodo A",       9),
            ("Nodo B",       9),
            ("X (p.u.)",    11),
            ("θ_A (°)",     11),
            ("θ_B (°)",     11),
            ("P_AB (p.u.)", 13),
            ("P_AB (MW)",   13),
            ("Sentido",     18),
        ]

        tabla_fl = tk.Frame(inner, bg=C['border'])
        tabla_fl.pack(fill=tk.X, padx=30, pady=(0, 20))

        head = tk.Frame(tabla_fl, bg=C['primary_d'])
        head.pack(fill=tk.X)
        for h, w in cols_fl:
            tk.Label(head, text=h, font=font(9, "bold"),
                     bg=C['primary_d'], fg='white',
                     width=w, pady=8).pack(side=tk.LEFT, padx=0)

        body = tk.Frame(tabla_fl, bg=C['border'])
        body.pack(fill=tk.X)
        for k, F in enumerate(flujos):
            bg_row = C['row_even'] if k % 2 == 0 else C['row_odd']
            P_pu = F['P']
            sentido = f"{F['a']+1} → {F['b']+1}" if P_pu >= 0 else f"{F['b']+1} → {F['a']+1}"
            color_P = C['success'] if P_pu >= 0 else C['danger']

            valores = [
                (f"{k+1}",                              bg_row, C['text'],  9),
                (f"{F['a']+1}",                         bg_row, C['text'],  9),
                (f"{F['b']+1}",                         bg_row, C['text'],  9),
                (f"{F['X']:.4f}",                       bg_row, C['text'], 11),
                (f"{math.degrees(theta[F['a']]):+.3f}", bg_row, C['primary'], 11),
                (f"{math.degrees(theta[F['b']]):+.3f}", bg_row, C['primary'], 11),
                (f"{P_pu:+.4f}",                        bg_row, color_P,   13),
                (f"{P_pu*BASE:+.2f}",                   bg_row, color_P,   13),
                (sentido,                               bg_row, color_P,   18),
            ]
            fila = tk.Frame(body, bg=C['border'])
            fila.pack(fill=tk.X, pady=1)
            for v, col_bg, col_fg, w in valores:
                tk.Label(fila, text=v, font=font_m(9),
                         bg=col_bg, fg=col_fg,
                         width=w, pady=6).pack(side=tk.LEFT, padx=0)

        # ============================================================
        # NOTA FINAL
        # ============================================================
        note = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        note.pack(fill=tk.X, padx=30, pady=(4, 20))
        tk.Label(note, text="NOTA",
                 font=font(9, "bold"), bg=C['surface'],
                 fg=C['text'], anchor='w').pack(fill=tk.X, padx=14, pady=(10, 2))
        tk.Label(note,
                 text="• El método DC aproxima la potencia activa linealmente:  "
                      "P_km = (θ_k − θ_m) / x_km.\n"
                      "• El nodo Slack balancea el sistema:  P_slack = − Σ P_neta,otros.\n"
                      "• Los valores negativos en θ indican ángulos en sentido "
                      "horario respecto al Slack (θ = 0).\n"
                      "• Un valor positivo de P_AB indica flujo del nodo A al nodo B; "
                      "negativo indica flujo en sentido contrario.\n"
                      "• Las pérdidas no se calculan en el modelo DC (se asume "
                      "x >> r). Los valores de Qd y Qg se muestran solo informativamente.",
                 font=font_m(9), bg=C['surface'], fg=C['text_l'],
                 justify='left', anchor='w', wraplength=1300
                 ).pack(fill=tk.X, padx=14, pady=(0, 12))

        tk.Button(inner, text="Cerrar ventana", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=30, pady=10, cursor='hand2',
                  borderwidth=0).pack(pady=(0, 24))

    def _titulo_seccion_resultados(self, parent, titulo, color):
        """Barra de título para secciones de resultados."""
        C = self.C
        f = tk.Frame(parent, bg=C['surface'], highlightthickness=1,
                     highlightbackground=C['border'])
        f.pack(fill=tk.X, padx=30, pady=(6, 10))
        tk.Frame(f, bg=color, width=5).pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(f, text=titulo, font=font(11, "bold"),
                 bg=C['surface'], fg=C['text'],
                 pady=8).pack(side=tk.LEFT, padx=14)
    # =================================================================
    # TEORÍA
    # =================================================================
    def mostrar_teoria(self):
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Teoría — Flujo DC Lineal")
        win.geometry("1000x760")
        win.configure(bg=C['bg'])

        hdr = tk.Frame(win, bg=C['primary'], height=70)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  TEORÍA — FLUJO DE POTENCIA LINEAL (DC)",
                 font=font(15, "bold"), bg=C['primary'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=18)
        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['amber'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=14, pady=6, cursor='hand2',
                  borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=14)

        canvas = tk.Canvas(win, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(win, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas.find_all()[0],
                                                 width=e.width))
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        def _wheel(e):
            canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        canvas.bind('<MouseWheel>', _wheel)
        inner.bind('<MouseWheel>', _wheel)

        def sec(txt):
            tk.Label(inner, text=txt, font=font(13, "bold"),
                     bg=C['bg'], fg=C['primary'], anchor='w'
                     ).pack(fill=tk.X, padx=30, pady=(22, 10))
        def par(txt):
            tk.Label(inner, text=txt, font=font(11),
                     bg=C['bg'], fg=C['text'], justify='left',
                     anchor='w', wraplength=900
                     ).pack(fill=tk.X, padx=34, pady=3)
        def form(txt):
            f = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                          highlightbackground=C['border'])
            f.pack(fill=tk.X, padx=50, pady=6)
            tk.Label(f, text=txt, font=font_m(11), bg=C['surface'],
                     fg=C['primary'], justify='left', anchor='w',
                     padx=16, pady=10).pack(fill=tk.X)

        sec("Principio del método DC")
        par("En sistemas de alta tensión, con redes predominantemente "
            "inductivas, es posible aproximar el flujo de potencia activa "
            "con un sistema de ecuaciones LINEALES, evitando iteraciones y "
            "reduciendo drásticamente el costo computacional.")

        sec("Supuestos del modelo")
        par("1. La reactancia de las líneas es mucho mayor que su resistencia: x_km >> r_km")
        par("2. Las tensiones nodales están cercanas a 1.0 p.u.: V_k ≈ V_m ≈ 1.0 p.u.")
        par("3. Las diferencias angulares entre nodos vecinos son pequeñas: "
            "cos(θ_km) ≈ 1  y  sin(θ_km) ≈ θ_km")

        sec("Ecuación matricial del flujo DC")
        form("ΔP = B · θ")
        par("Donde ΔP es el vector de inyecciones netas de potencia activa "
            "(Pg − Pd), B es la matriz de susceptancias y θ es el vector de "
            "ángulos nodales.")
        form("θ = F · ΔP      con      F = B⁻¹")

        sec("Elementos de la matriz B")
        par("Con el nodo Slack seleccionado (θ_slack = 0):")
        form("B_km = − 1 / x_km           (fuera de la diagonal)")
        form("B_kk = Σₘ 1 / x_km          (diagonal)")

        sec("Flujo por cada línea")
        form("P_km = (θ_k − θ_m) / x_km")
        par("El signo de P_km indica la dirección del flujo: positivo → "
            "fluye de k a m; negativo → fluye de m a k.")

        sec("Ventajas y limitaciones")
        par("• Ventajas:  sin iteraciones, muy rápido, adecuado para estudios "
            "de contingencias N-1 y análisis de mercados eléctricos.")
        par("• Limitaciones:  no considera pérdidas, no resuelve tensiones, "
            "solo potencias activas. Los resultados son aproximados.")

        tk.Button(inner, text="Cerrar", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=32, pady=10, cursor='hand2',
                  borderwidth=0).pack(pady=30)

    # =================================================================
    # VOLVER AL MENÚ
    # =================================================================
    def _volver_al_menu(self):
        if not messagebox.askyesno("Confirmar",
            "¿Volver al menú principal?\nSe perderán los datos no guardados."):
            return
        for w in self.root.winfo_children():
            w.destroy()
        try:
            def abrir_simulador(root):
                root.title("Vatia — Flujos de Potencia (Newton-Raphson)")
                root.geometry("1500x900")
                root.minsize(1180, 720)
                root.configure(bg="#EDF1F6")
                app = AppFlujoPotencia(root)
                root.app = app
            DashboardScreen(self.root, abrir_simulador)
        except Exception:
            import traceback
            traceback.print_exc()
            
# =====================================================================
# MÓDULO — ANÁLISIS DE FALLAS (ENTRADA DE R, X, B/2 Y SHUNTS EN MVAR)
# =====================================================================
class FaultAnalysisScreen:
    """Módulo de análisis de fallas con componentes simétricas."""

    C = {
        'bg':          '#EDF1F6', 'panel':       '#FFFFFF', 'surface':     '#F7F9FC',
        'sidebar':     '#E2E8F0', 'primary':     '#1E3A5F', 'primary_l':   '#2C5282',
        'primary_d':   '#152741', 'accent':      '#0E7490', 'accent_l':    '#0891B2',
        'slack':       '#047857', 'pv':          '#6D28D9', 'pq':          '#C2410C',
        'info':        '#0369A1', 'amber':       '#B45309', 'red':         '#B91C1C',
        'text':        '#1F2937', 'text_l':      '#4B5563', 'text_ll':     '#94A3B8',
        'border':      '#CBD5E1', 'head_bg':     '#E2E8F0', 'row_even':    '#FFFFFF',
        'row_odd':     '#F7F9FC', 'success':     '#047857', 'warning':     '#B45309',
        'danger':      '#B91C1C', 'gray_dark':   '#475569',
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Vatia — Análisis de Fallas")
        self.root.geometry("1500x900")
        self.root.minsize(1180, 720)
        self.root.configure(bg=self.C['bg'])

        self.resultados = None
        self.BASE_MVA = 100.0
        self.V_BASE_KV = 220.0

        self.tipo_barras = []
        self.V_pref_widgets = []
        self.ang_pref_widgets = []
        self.entries_lineas = []
        self.entries_trafos = []
        self.entries_inductores = []
        self.entries_capacitores = []
        self.entries_cargas = [] 
        self.gen_widgets = []
        self.sidebar_labels = {}

        self.combo_barra_falla = None
        self.combo_tipo_falla = None
        self.entry_ZF = None
        self.entry_Zg = None

        self._updating = False
        self._sidebar_after_id = None

        self._build_menu()
        self._build_header()
        self._build_toolbar()
        self._build_base_bar()
        self._build_main_area()
        self._build_action_bar()
        self._build_status_bar()

        self.root.bind('<Control-n>', lambda e: self._nuevo_proyecto())
        self.root.bind('<Control-o>', lambda e: self._abrir_proyecto())
        self.root.bind('<Control-s>', lambda e: self._guardar_proyecto())
        self.root.bind('<Control-e>', lambda e: self._exportar_excel())
        self.root.bind('<F5>', lambda e: self.calcular_falla())
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.generar_tablas()

    # =================================================================
    # MENÚ Y HEADER
    # =================================================================
    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        m_arch = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Archivo", menu=m_arch)
        m_arch.add_command(label="Nuevo proyecto\t\tCtrl+N", command=self._nuevo_proyecto)
        m_arch.add_command(label="Abrir proyecto...\tCtrl+O", command=self._abrir_proyecto)
        m_arch.add_command(label="Guardar\t\tCtrl+S", command=self._guardar_proyecto)
        m_arch.add_separator()
        m_arch.add_command(label="Exportar a Excel...\tCtrl+E", command=self._exportar_excel)
        m_arch.add_separator()
        m_arch.add_command(label="Volver al menú", command=self._volver_al_menu)
        m_calc = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Cálculo", menu=m_calc)
        m_calc.add_command(label="Calcular falla\tF5", command=self.calcular_falla)
        m_calc.add_separator()
        m_calc.add_command(label="Ver resultados", command=self.mostrar_resultados)
        m_calc.add_command(label="Procedimiento paso a paso", command=self.mostrar_procedimiento)
        m_ayuda = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=m_ayuda)
        m_ayuda.add_command(label="Teoría de fallas", command=self.mostrar_teoria)

    def _build_header(self):
        C = self.C
        hdr = tk.Frame(self.root, bg=C['primary_d'], height=90)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['red'], height=4).pack(fill=tk.X, side=tk.TOP)
        left = tk.Frame(hdr, bg=C['primary_d'])
        left.pack(side=tk.LEFT, padx=24, pady=14)
        cv = tk.Canvas(left, width=54, height=54, bg=C['primary_d'], highlightthickness=0)
        cv.pack(side=tk.LEFT, padx=(0, 14))
        cv.create_oval(2, 2, 52, 52, fill=C['red'], outline='')
        cv.create_polygon(30, 10, 22, 30, 28, 30, 25, 46, 35, 25, 29, 25, 33, 10, fill='white', outline='')
        tinfo = tk.Frame(left, bg=C['primary_d']); tinfo.pack(side=tk.LEFT)
        tk.Label(tinfo, text="ANÁLISIS DE FALLAS ELÉCTRICAS", font=font(15, "bold"), bg=C['primary_d'], fg='white').pack(anchor='w')
        tk.Label(tinfo, text="Componentes simétricas  ·  SLG  ·  LL  ·  LLG  ·  LLL", font=font(9, italic=True), bg=C['primary_d'], fg=C['text_l']).pack(anchor='w')
        right = tk.Frame(hdr, bg=C['primary_d']); right.pack(side=tk.RIGHT, padx=24, pady=15)
        badge = tk.Frame(right, bg=C['red']); badge.pack()
        tk.Label(badge, text="  MÉTODO SIMÉTRICO  ", font=font(10, "bold"), bg=C['red'], fg='white', padx=10, pady=7).pack()

    # =================================================================
    # TOOLBAR Y BASES
    # =================================================================
    def _build_toolbar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(10, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1, highlightbackground=C['border']); card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel']); inner.pack(padx=18, pady=12, fill=tk.X)

        tk.Label(inner, text="Barras:", font=font(10, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT)
        self.entry_n = ttk.Entry(inner, width=3, justify='center', font=font(11)); self.entry_n.pack(side=tk.LEFT, padx=5, ipady=3); self.entry_n.insert(0, "7")

        tk.Label(inner, text="Líneas:", font=font(10, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(10, 0))
        self.entry_n_lineas = ttk.Entry(inner, width=3, justify='center', font=font(11)); self.entry_n_lineas.pack(side=tk.LEFT, padx=5, ipady=3); self.entry_n_lineas.insert(0, "7")

        tk.Label(inner, text="Trafos:", font=font(10, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(10, 0))
        self.entry_n_trafos = ttk.Entry(inner, width=3, justify='center', font=font(11)); self.entry_n_trafos.pack(side=tk.LEFT, padx=5, ipady=3); self.entry_n_trafos.insert(0, "3")

        tk.Label(inner, text="Ind:", font=font(10, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(10, 0))
        self.entry_n_ind = ttk.Entry(inner, width=3, justify='center', font=font(11)); self.entry_n_ind.pack(side=tk.LEFT, padx=5, ipady=3); self.entry_n_ind.insert(0, "1")

        tk.Label(inner, text="Cap:", font=font(10, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(10, 0))
        self.entry_n_cap = ttk.Entry(inner, width=3, justify='center', font=font(11)); self.entry_n_cap.pack(side=tk.LEFT, padx=5, ipady=3); self.entry_n_cap.insert(0, "2")

        tk.Label(inner, text="Gen:", font=font(10, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(10, 0))
        self.entry_n_gen = ttk.Entry(inner, width=3, justify='center', font=font(11)); self.entry_n_gen.pack(side=tk.LEFT, padx=5, ipady=3); self.entry_n_gen.insert(0, "2")

        tk.Label(inner, text="Cargas Z:", font=font(10, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(10, 0))
        self.entry_n_cargas = ttk.Entry(inner, width=3, justify='center', font=font(11))
        self.entry_n_cargas.pack(side=tk.LEFT, padx=5, ipady=3)
        self.entry_n_cargas.insert(0, "1")

        tk.Button(inner, text="  Generar Tablas  ", command=self.generar_tablas, bg=C['primary'], fg='white', font=font(10, "bold"), relief=tk.FLAT, cursor='hand2', padx=16, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=10)

        tk.Button(inner, text="  Generar Tablas  ", command=self.generar_tablas, bg=C['primary'], fg='white', font=font(10, "bold"), relief=tk.FLAT, cursor='hand2', padx=16, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=10)
        tk.Button(inner, text="◄  Volver al menú", command=self._volver_al_menu, bg=C['surface'], fg=C['primary'], font=font(10, "bold"), relief=tk.FLAT, cursor='hand2', padx=16, pady=7, borderwidth=0).pack(side=tk.RIGHT, padx=2)

    def _build_base_bar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg']); outer.pack(fill=tk.X, padx=14, pady=(0, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1, highlightbackground=C['border']); card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel']); inner.pack(padx=18, pady=10, fill=tk.X)

        tk.Label(inner, text="BASES DEL SISTEMA", font=font(9, "bold"), bg=C['panel'], fg=C['primary']).pack(side=tk.LEFT, padx=(0, 20))
        tk.Label(inner, text="Potencia base:", font=font(9), bg=C['panel'], fg=C['text_l']).pack(side=tk.LEFT)
        self.entry_base_mva = ttk.Entry(inner, width=8, justify='center', font=font(10)); self.entry_base_mva.pack(side=tk.LEFT, padx=(6, 4), ipady=2); self.entry_base_mva.insert(0, "100")
        tk.Label(inner, text="MVA", font=font(9, "bold"), bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=(0, 24))

        tk.Label(inner, text="Voltaje base:", font=font(9), bg=C['panel'], fg=C['text_l']).pack(side=tk.LEFT)
        self.entry_base_kv = ttk.Entry(inner, width=8, justify='center', font=font(10)); self.entry_base_kv.pack(side=tk.LEFT, padx=(6, 4), ipady=2); self.entry_base_kv.insert(0, "220")
        tk.Label(inner, text="kV", font=font(9, "bold"), bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=(0, 24))

        self.lbl_bases_info = tk.Label(inner, text="→  Z_base = (V_base)² / S_base    ·    I_base = S_base / (√3 · V_base)", font=font_m(9, "italic"), bg=C['panel'], fg=C['text_ll']); self.lbl_bases_info.pack(side=tk.LEFT)
        self.entry_base_mva.bind('<KeyRelease>', lambda e: self._actualizar_bases())
        self.entry_base_kv.bind('<KeyRelease>', lambda e: self._actualizar_bases())
        tk.Label(inner, text="kV", font=font(9, "bold"), bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=(0, 24))

        # === NUEVO: Agregar entrada para V Prefalla Global ===
        tk.Label(inner, text="V_prefalla global:", font=font(9), bg=C['panel'], fg=C['text_l']).pack(side=tk.LEFT)
        self.entry_v_pref_global = ttk.Entry(inner, width=6, justify='center', font=font(10))
        self.entry_v_pref_global.pack(side=tk.LEFT, padx=(6, 4), ipady=2)
        self.entry_v_pref_global.insert(0, "1.0")
        
        btn_apply_v = tk.Button(inner, text="Aplicar a todas las barras", command=self._aplicar_v_pref_global,
                                bg=C['accent'], fg='white', font=font(8, "bold"), relief=tk.FLAT, 
                                cursor='hand2', padx=8, pady=3, borderwidth=0)
        btn_apply_v.pack(side=tk.LEFT, padx=(0, 24))
        # ====================================================

        self.lbl_bases_info = tk.Label(inner, text="→  Z_base = (V_base)² / S_base    ·    I_base = S_base / (√3 · V_base)", font=font_m(9, "italic"), bg=C['panel'], fg=C['text_ll'])
        self.lbl_bases_info.pack(side=tk.LEFT)

    def _actualizar_bases(self):
        try:
            self.BASE_MVA = self._parse_float(self.entry_base_mva.get(), 100.0)
            if self.BASE_MVA <= 0: self.BASE_MVA = 100.0
        except Exception: pass
        try:
            self.V_BASE_KV = self._parse_float(self.entry_base_kv.get(), 220.0)
            if self.V_BASE_KV <= 0: self.V_BASE_KV = 220.0
        except Exception: pass
        try:
            z_b = self.V_BASE_KV ** 2 / self.BASE_MVA
            i_b = self.BASE_MVA / (math.sqrt(3) * self.V_BASE_KV)
            self.lbl_bases_info.config(text=f"→  Z_base = {z_b:.4f} Ω    ·    I_base = {i_b:.4f} kA")
        except Exception: pass
    def _aplicar_v_pref_global(self):
        """Aplica el voltaje prefalla global a todas las barras de la pestaña."""
        try:
            v = self._parse_float(self.entry_v_pref_global.get(), 1.0)
            for e in self.V_pref_widgets:
                e.delete(0, tk.END)
                e.insert(0, f"{v:.4f}")
            self._set_status(f"Voltaje prefalla de {v} p.u. aplicado a todas las barras.", 'ok')
        except Exception:
            pass

    # =================================================================
    # MAIN AREA
    # =================================================================
    def _build_main_area(self):
        C = self.C
        main = tk.Frame(self.root, bg=C['bg']); main.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        self.paned = ttk.PanedWindow(main, orient=tk.HORIZONTAL); self.paned.pack(fill=tk.BOTH, expand=True)
        self.sidebar = tk.Frame(self.paned, bg=C['sidebar'], width=210); self.paned.add(self.sidebar, weight=0); self._build_sidebar()
        center = tk.Frame(self.paned, bg=C['bg']); self.paned.add(center, weight=5)
        self.notebook = ttk.Notebook(center); self.notebook.pack(fill=tk.BOTH, expand=True)
        right = tk.Frame(self.paned, bg=C['bg'], width=380); self.paned.add(right, weight=4)
        right_lf = tk.Frame(right, bg=C['panel'], highlightthickness=1, highlightbackground=C['border']); right_lf.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        tk.Label(right_lf, text="  VISTA PREVIA DEL SISTEMA  ", font=font(10, "bold"), bg=C['head_bg'], fg=C['red'], pady=8).pack(fill=tk.X)
        self.preview_canvas = tk.Canvas(right_lf, bg=C['panel'], highlightthickness=0); self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.preview_canvas.bind('<Configure>', lambda e: self._dibujar_preview())

    def _build_sidebar(self):
        C = self.C
        sb = self.sidebar; sb.pack_propagate(False)
        tk.Label(sb, text="  PANEL DE CONTROL  ", font=font(10, "bold"), bg=C['red'], fg='white', pady=10).pack(fill=tk.X)
        tk.Frame(sb, bg=C['accent'], height=2).pack(fill=tk.X)
        def _stat(key, label, color=None):
            frame = tk.Frame(sb, bg=C['sidebar']); frame.pack(fill=tk.X, padx=14, pady=(8, 2))
            tk.Label(frame, text=label.upper(), font=font(8, "bold"), bg=C['sidebar'], fg=C['text_l']).pack(anchor='w')
            lbl = tk.Label(frame, text="—", font=font(13, "bold"), bg=C['sidebar'], fg=color or C['text'], anchor='w', justify='left', wraplength=180)
            lbl.pack(anchor='w', fill=tk.X); self.sidebar_labels[key] = lbl
        def _sep(): tk.Frame(sb, bg=C['border'], height=1).pack(fill=tk.X, padx=14, pady=6)
        _stat('barras', 'Barras', C['accent']); _stat('lineas', 'Líneas', C['info'])
        _stat('trafos', 'Transformadores', C['primary_l']); _stat('shunts', 'Shunts (Ind/Cap)', C['amber'])
        _stat('generadores', 'Generadores', C['pv']); _sep()
        _stat('barra_falla', 'Barra fallada', C['red']); _stat('tipo_falla', 'Tipo de falla', C['red']); _sep()
        _stat('estado', 'Estado', C['text_l'])
        pie = tk.Frame(sb, bg=C['sidebar']); pie.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        tk.Label(pie, text="por Kevin", font=font(9, italic=True), bg=C['sidebar'], fg=C['text_l']).pack()
        tk.Label(pie, text="Vatia © 2026", font=font_m(8), bg=C['sidebar'], fg=C['text_ll']).pack()
        self.root.after(400, self._update_sidebar)

    def _update_sidebar(self):
        try:
            if not self.root.winfo_exists(): return
        except Exception: return
        if not getattr(self, 'sidebar_labels', None):
            try: self._sidebar_after_id = self.root.after(500, self._update_sidebar)
            except Exception: pass
            return
        try:
            if not self.entry_n.winfo_exists(): return
        except Exception: return

        try:
            n = int(self.entry_n.get())
            n_lineas = sum(1 for row in self.entries_lineas if self._parse_int(row[0].get(), 0) > 0 and self._parse_int(row[1].get(), 0) > 0)
            n_trafos = sum(1 for row in self.entries_trafos if self._parse_int(row[0].get(), 0) > 0 and self._parse_int(row[1].get(), 0) > 0)
            n_ind = sum(1 for row in self.entries_inductores if self._parse_int(row[0].get(), 0) > 0)
            n_cap = sum(1 for row in self.entries_capacitores if self._parse_int(row[0].get(), 0) > 0)
            n_cargas = sum(1 for row in self.entries_cargas if self._parse_int(row[0].get(), 0) > 0) 
            n_gen = sum(1 for g in self.gen_widgets if self._parse_float(g[3].get(), 0.0) > 1e-9)

            self.sidebar_labels['barras'].config(text=str(n))
            self.sidebar_labels['lineas'].config(text=str(n_lineas))
            self.sidebar_labels['trafos'].config(text=str(n_trafos))
            self.sidebar_labels['shunts'].config(text=f"{n_ind} / {n_cap} / {n_cargas}") 
            self.sidebar_labels['generadores'].config(text=str(n_gen))

            if self.combo_barra_falla and self.combo_barra_falla.winfo_exists():
                self.sidebar_labels['barra_falla'].config(text=self.combo_barra_falla.get() or "—")
            if self.combo_tipo_falla and self.combo_tipo_falla.winfo_exists():
                tf = self.combo_tipo_falla.get()
                tf_corto = (tf.replace("Monofásica", "Monof.").replace("Línea-Línea", "L-L").replace("Doble Línea-Tierra", "Doble L-T").replace("Trifásica", "Trif."))
                self.sidebar_labels['tipo_falla'].config(text=tf_corto or "—")

            if self.resultados: self.sidebar_labels['estado'].config(text="CALCULADO", fg=self.C['success'])
            else: self.sidebar_labels['estado'].config(text="SIN CALCULAR", fg=self.C['text_l'])
        except tk.TclError: return
        except Exception: pass
        try:
            if self.root.winfo_exists(): self._sidebar_after_id = self.root.after(500, self._update_sidebar)
        except Exception: pass

    # =================================================================
    # ACTION BAR Y STATUS
    # =================================================================
    def _build_action_bar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg']); outer.pack(fill=tk.X, padx=14, pady=(6, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1, highlightbackground=C['border']); card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel']); inner.pack(padx=14, pady=10, fill=tk.X)
        def _grupo(parent, titulo):
            g = tk.Frame(parent, bg=C['panel']); g.pack(side=tk.LEFT, padx=(0, 20))
            tk.Label(g, text=titulo.upper(), font=font(8, "bold"), bg=C['panel'], fg=C['text_l']).pack(anchor='w', pady=(0, 4))
            fila = tk.Frame(g, bg=C['panel']); fila.pack(); return fila
        g1 = _grupo(inner, "Cálculo")
        btn_calc = tk.Button(g1, text="  ⚡  Calcular Falla  ", command=self.calcular_falla, bg=C['red'], fg='white', font=font(11, "bold"), relief=tk.FLAT, cursor='hand2', padx=20, pady=7, borderwidth=0); btn_calc.pack(side=tk.LEFT, padx=2)
        g2 = _grupo(inner, "Vista")
        btn_res = tk.Button(g2, text="  Ver Resultados  ", command=self.mostrar_resultados, bg=C['accent'], fg='white', font=font(10, "bold"), relief=tk.FLAT, cursor='hand2', padx=16, pady=7, borderwidth=0); btn_res.pack(side=tk.LEFT, padx=2)
        btn_proc = tk.Button(g2, text="  Procedimiento  ", command=self.mostrar_procedimiento, bg=C['info'], fg='white', font=font(10, "bold"), relief=tk.FLAT, cursor='hand2', padx=16, pady=7, borderwidth=0); btn_proc.pack(side=tk.LEFT, padx=2)
        g3 = tk.Frame(inner, bg=C['panel']); g3.pack(side=tk.RIGHT)
        tk.Label(g3, text="ARCHIVO / SOPORTE", font=font(8, "bold"), bg=C['panel'], fg=C['text_l']).pack(anchor='e', pady=(0, 4))
        fila3 = tk.Frame(g3, bg=C['panel']); fila3.pack()
        btn_exp = tk.Button(fila3, text="  Exportar Excel  ", command=self._exportar_excel, bg=C['primary'], fg='white', font=font(10, "bold"), relief=tk.FLAT, cursor='hand2', padx=16, pady=7, borderwidth=0); btn_exp.pack(side=tk.LEFT, padx=2)
        btn_teo = tk.Button(fila3, text="  Teoría  ", command=self.mostrar_teoria, bg=C['surface'], fg=C['primary'], font=font(10, "bold"), relief=tk.FLAT, cursor='hand2', padx=16, pady=7, borderwidth=0); btn_teo.pack(side=tk.LEFT, padx=2)

    def _build_status_bar(self):
        C = self.C
        sb = tk.Frame(self.root, bg=C['primary_d'], height=32); sb.pack(fill=tk.X, side=tk.BOTTOM); sb.pack_propagate(False)
        tk.Frame(sb, bg=C['red'], width=3).pack(side=tk.LEFT, fill=tk.Y)
        self.status_icon = tk.Label(sb, text="●", font=font(11), bg=C['primary_d'], fg=C['success']); self.status_icon.pack(side=tk.LEFT, padx=(12, 4))
        self.status_var = tk.StringVar(value="Listo.")
        tk.Label(sb, textvariable=self.status_var, font=font(9), bg=C['primary_d'], fg='white').pack(side=tk.LEFT)
        tk.Label(sb, text="v1.0  ·  Módulo de Fallas", font=font_m(8), bg=C['primary_d'], fg=C['text_ll']).pack(side=tk.RIGHT, padx=(0, 14))

    def _set_status(self, text, kind='info'):
        try:
            self.status_var.set(text)
            colors = {'info': self.C['accent'], 'ok': self.C['success'], 'warn': self.C['warning'], 'err': self.C['danger']}
            self.status_icon.config(fg=colors.get(kind, self.C['accent']))
        except Exception: pass

    # =================================================================
    # HELPERS
    # =================================================================
    @staticmethod
    def _parse_float(s, default=0.0):
        try:
            if s is None: return default
            s = str(s).strip().replace(',', '.')
            return float(s) if s else default
        except Exception: return default
    @staticmethod
    def _parse_int(s, default=0):
        try:
            if s is None: return default
            s = str(s).strip().replace(',', '.')
            return int(float(s)) if s else default
        except Exception: return default

    def _crear_tab_scrolleable(self, parent):
        canvas = tk.Canvas(parent, bg=self.C['panel'], highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        interior = tk.Frame(canvas, bg=self.C['panel'])
        interior.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=interior, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas.find_all()[0], width=e.width))
        def _on_mousewheel(event): canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        def _bind_wheel(e): canvas.bind_all('<MouseWheel>', _on_mousewheel)
        def _unbind_wheel(e): canvas.unbind_all('<MouseWheel>')
        canvas.bind('<Enter>', _bind_wheel); canvas.bind('<Leave>', _unbind_wheel)
        return interior

    # =================================================================
    # GENERAR TABLAS
    # =================================================================
    def generar_tablas(self):
        try:
            n = int(self.entry_n.get()); n_lineas = int(self.entry_n_lineas.get())
            n_trafos = int(self.entry_n_trafos.get()); n_ind = int(self.entry_n_ind.get())
            n_cap = int(self.entry_n_cap.get()); n_gen = int(self.entry_n_gen.get())
            n_cargas = int(self.entry_n_cargas.get()) # <--- AGREGAR
        except ValueError:
            messagebox.showwarning("Advertencia", "Ingrese números enteros válidos."); return

        for tab_id in self.notebook.tabs(): self.notebook.forget(tab_id)
        self.tipo_barras.clear(); self.V_pref_widgets.clear(); self.ang_pref_widgets.clear()
        self.entries_lineas.clear(); self.entries_trafos.clear(); self.entries_inductores.clear()
        self.entries_capacitores.clear(); self.entries_cargas.clear(); # <--- AGREGAR
        self.gen_widgets.clear(); self.resultados = None

        self.tab_barras = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_lineas = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_trafos = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_ind = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_cap = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_cargas = tk.Frame(self.notebook, bg=self.C['panel']) # <--- AGREGAR
        self.tab_gen = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_falla = tk.Frame(self.notebook, bg=self.C['panel'])

        self.notebook.add(self.tab_barras, text='   Barras   ')
        self.notebook.add(self.tab_lineas, text='   Líneas   ')
        self.notebook.add(self.tab_trafos, text='   Transformadores   ')
        self.notebook.add(self.tab_ind, text='   Inductores   ')
        self.notebook.add(self.tab_cap, text='   Capacitores   ')
        self.notebook.add(self.tab_cargas, text='   Cargas (Z cte)   ') # <--- AGREGAR
        self.notebook.add(self.tab_gen, text='   Generadores   ')
        self.notebook.add(self.tab_falla, text='   Falla   ')

        self._crear_tab_barras(n); self._crear_tab_lineas(n_lineas); self._crear_tab_trafos(n_trafos)
        self._crear_tab_inductores(n_ind); self._crear_tab_capacitores(n_cap)
        self._crear_tab_cargas(n_cargas) # <--- AGREGAR
        self._crear_tab_generadores(n_gen); self._crear_tab_falla(n)

        self.tipo_barras[0].set("Slack (ref.)")
        for i in range(1, n): self.tipo_barras[i].set("PQ")
        self.V_pref_widgets[0].delete(0, tk.END); self.V_pref_widgets[0].insert(0, "1.0")

        self._set_status(f"Sistema: {n} barras, {n_lineas} líneas, {n_trafos} trafos, {n_ind} ind, {n_cap} cap, {n_cargas} cargas, {n_gen} gen.", 'info')
        self._update_sidebar(); self.root.after(120, self._dibujar_preview)

    def _crear_tab_barras(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_barras)
        hdr = tk.Frame(tab, bg=C['primary_d']); hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS DE LAS BARRAS", font=font(12, "bold"), bg=C['primary_d'], fg='white', pady=12).pack()
        grid = tk.Frame(tab, bg=C['panel']); grid.pack(padx=20, pady=20, fill=tk.X)
        headers = ["Barra", "Tipo", "V_pref (p.u.)", "Ángulo (°)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"), bg=C['head_bg'], fg=C['accent'], padx=8, pady=7).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)
        for i in range(n):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']
            tk.Label(grid, text=str(i+1), font=font(10, "bold"), bg=bg, fg=C['red'], padx=8, pady=6).grid(row=i+1, column=0, padx=1, pady=1, sticky='nsew')
            cb = ttk.Combobox(grid, values=["Slack (ref.)", "PV", "PQ"], state="readonly", font=font(9))
            cb.grid(row=i+1, column=1, padx=1, pady=1, sticky='nsew'); cb.set("PQ"); cb.bind("<<ComboboxSelected>>", lambda e, idx=i: self._on_tipo_changed(idx))
            self.tipo_barras.append(cb)
            e_v = ttk.Entry(grid, justify='center', font=font(9)); e_v.grid(row=i+1, column=2, padx=1, pady=1, ipady=4, sticky='nsew'); e_v.insert(0, "1.0")
            self.V_pref_widgets.append(e_v)
            e_a = ttk.Entry(grid, justify='center', font=font(9)); e_a.grid(row=i+1, column=3, padx=1, pady=1, ipady=4, sticky='nsew'); e_a.insert(0, "0.0")
            self.ang_pref_widgets.append(e_a)

    def _crear_tab_lineas(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_lineas)
        hdr = tk.Frame(tab, bg=C['primary_d']); hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS DE LAS LÍNEAS (R, X, B/2)", font=font(12, "bold"), bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Ingrese R, X en p.u., B/2 (susceptancia) y Z0 (R0, X0). Si Z0=0, se asume 3*Z1.", font=font(9, italic=True), bg=C['primary_d'], fg=C['text_l']).pack(pady=(0, 10))
        grid = tk.Frame(tab, bg=C['panel']); grid.pack(padx=20, pady=20, fill=tk.X)
        headers = ["Barra A", "Barra B", "R₁", "X₁", "B₁/2", "R₀", "X₀"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"), bg=C['head_bg'], fg=C['accent'], padx=4, pady=7).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)
        for i in range(n):
            row = []
            for j in range(7):
                e = ttk.Entry(grid, justify='center', font=font(9)); e.grid(row=i+1, column=j, padx=1, pady=1, ipady=4, sticky='nsew')
                if j == 2: e.insert(0, "0.01")
                elif j == 3: e.insert(0, "0.1")
                elif j == 4: e.insert(0, "0.025")
                elif j == 5: e.insert(0, "0.03")
                elif j == 6: e.insert(0, "0.3")
                e.bind('<KeyRelease>', lambda ev: (self._dibujar_preview(), self._update_sidebar()))
                e.bind('<FocusOut>', lambda ev: (self._dibujar_preview(), self._update_sidebar()))
                row.append(e)
            self.entries_lineas.append(row)

    def _crear_tab_trafos(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_trafos)
        hdr = tk.Frame(tab, bg=C['primary_d']); hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS DE LOS TRANSFORMADORES (R, X)", font=font(12, "bold"), bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Ingrese R, X (Z1=Z2=Z0) y seleccione la configuración.", font=font(9, italic=True), bg=C['primary_d'], fg=C['text_l']).pack(pady=(0, 10))
        grid = tk.Frame(tab, bg=C['panel']); grid.pack(padx=20, pady=20, fill=tk.X)
        headers = ["Barra A", "Barra B", "R", "X", "Conexión"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"), bg=C['head_bg'], fg=C['accent'], padx=8, pady=7).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)
        configs = ["YNyn (Y aterrizado - Y aterrizado)", "YNd (Y aterrizado - Delta)", "Dyn (Delta - Y aterrizado)", "Yy (Y sin aterrizar - Y sin aterrizar)", "Dd (Delta - Delta)"]
        for i in range(n):
            row = []
            for j in range(5):
                if j == 4:
                    cb = ttk.Combobox(grid, values=configs, state="readonly", font=font(9)); cb.grid(row=i+1, column=j, padx=1, pady=1, sticky='nsew'); cb.set(configs[0]); cb.bind("<<ComboboxSelected>>", lambda ev: self._update_sidebar())
                    row.append(cb)
                else:
                    e = ttk.Entry(grid, justify='center', font=font(9)); e.grid(row=i+1, column=j, padx=1, pady=1, ipady=4, sticky='nsew')
                    if j == 2: e.insert(0, "0.01")
                    elif j == 3: e.insert(0, "0.1")
                    e.bind('<KeyRelease>', lambda ev: (self._dibujar_preview(), self._update_sidebar()))
                    e.bind('<FocusOut>', lambda ev: (self._dibujar_preview(), self._update_sidebar()))
                    row.append(e)
            self.entries_trafos.append(row)

    def _crear_tab_inductores(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_ind)
        hdr = tk.Frame(tab, bg=C['primary_d']); hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS DE LOS INDUCTORES (SHUNTS)", font=font(12, "bold"), bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Ingrese la barra y la potencia reactiva absorbida Q en MVAr a tensión nominal.", font=font(9, italic=True), bg=C['primary_d'], fg=C['text_l']).pack(pady=(0, 10))
        grid = tk.Frame(tab, bg=C['panel']); grid.pack(padx=20, pady=20, fill=tk.X)
        headers = ["Nodo", "Q_abs (MVAr)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"), bg=C['head_bg'], fg=C['accent'], padx=8, pady=7).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)
        for i in range(n):
            row = []
            for j in range(2):
                e = ttk.Entry(grid, justify='center', font=font(9)); e.grid(row=i+1, column=j, padx=1, pady=1, ipady=4, sticky='nsew')
                if j == 0: e.insert(0, "1")
                elif j == 1: e.insert(0, "50.0")
                e.bind('<KeyRelease>', lambda ev: self._update_sidebar())
                row.append(e)
            self.entries_inductores.append(row)

    def _crear_tab_capacitores(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_cap)
        hdr = tk.Frame(tab, bg=C['primary_d']); hdr.pack(fill=tk.X)
        tk.Label(hdr, text="DATOS DE LOS CAPACITORES (SHUNTS)", font=font(12, "bold"), bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Ingrese la barra y la potencia reactiva entregada Q en MVAr a tensión nominal.", font=font(9, italic=True), bg=C['primary_d'], fg=C['text_l']).pack(pady=(0, 10))
        grid = tk.Frame(tab, bg=C['panel']); grid.pack(padx=20, pady=20, fill=tk.X)
        headers = ["Nodo", "Q_ent (MVAr)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"), bg=C['head_bg'], fg=C['accent'], padx=8, pady=7).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)
        for i in range(n):
            row = []
            for j in range(2):
                e = ttk.Entry(grid, justify='center', font=font(9)); e.grid(row=i+1, column=j, padx=1, pady=1, ipady=4, sticky='nsew')
                if j == 0: e.insert(0, "1")
                elif j == 1: e.insert(0, "50.0")
                e.bind('<KeyRelease>', lambda ev: self._update_sidebar())
                row.append(e)
            self.entries_capacitores.append(row)

    def _crear_tab_cargas(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_cargas)
        hdr = tk.Frame(tab, bg=C['primary_d']); hdr.pack(fill=tk.X)
        tk.Label(hdr, text="CARGAS DE IMPEDANCIA CONSTANTE", font=font(12, "bold"), bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Ingrese la barra, P y Q de la carga (se modelará como admitancia shunt Y = P - jQ).", font=font(9, italic=True), bg=C['primary_d'], fg=C['text_l']).pack(pady=(0, 10))
        grid = tk.Frame(tab, bg=C['panel']); grid.pack(padx=20, pady=20, fill=tk.X)
        headers = ["Nodo", "P (MW)", "Q (MVAr)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"), bg=C['head_bg'], fg=C['accent'], padx=8, pady=7).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)
        for i in range(n):
            row = []
            for j in range(3):
                e = ttk.Entry(grid, justify='center', font=font(9)); e.grid(row=i+1, column=j, padx=1, pady=1, ipady=4, sticky='nsew')
                if j == 0: e.insert(0, "1")
                else: e.insert(0, "0.0")
                e.bind('<KeyRelease>', lambda ev: self._update_sidebar())
                row.append(e)
            self.entries_cargas.append(row)

    def _crear_tab_generadores(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_gen)
        hdr = tk.Frame(tab, bg=C['primary_d']); hdr.pack(fill=tk.X)
        tk.Label(hdr, text="IMPEDANCIAS DE LOS GENERADORES", font=font(12, "bold"), bg=C['primary_d'], fg='white', pady=12).pack()
        tk.Label(hdr, text="Seleccione el nodo e ingrese R y X para cada secuencia.", font=font(9, italic=True), bg=C['primary_d'], fg=C['text_l']).pack(pady=(0, 10))
        grid = tk.Frame(tab, bg=C['panel']); grid.pack(padx=20, pady=20, fill=tk.X)
        headers = ["Gen #", "Nodo", "R₁", "X₁", "R₂", "X₂", "R₀", "X₀"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"), bg=C['head_bg'], fg=C['accent'], padx=4, pady=7).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)
        for i in range(n):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']
            tk.Label(grid, text=f"G{i+1}", font=font(10, "bold"), bg=bg, fg=C['pv'], padx=4, pady=6).grid(row=i+1, column=0, padx=1, pady=1, sticky='nsew')
            row = []
            for j in range(7):
                e = ttk.Entry(grid, justify='center', font=font(9)); e.grid(row=i+1, column=j+1, padx=1, pady=1, ipady=4, sticky='nsew')
                if j == 0: e.insert(0, "1")
                elif j == 1 or j == 3: e.insert(0, "0.0")
                elif j == 2 or j == 5: e.insert(0, "0.1")
                elif j == 4 or j == 6: e.insert(0, "0.05")
                e.bind('<KeyRelease>', lambda ev: self._update_sidebar())
                row.append(e)
            self.gen_widgets.append(row)

    def _crear_tab_falla(self, n):
        C = self.C
        tab = self._crear_tab_scrolleable(self.tab_falla)
        hdr = tk.Frame(tab, bg=C['red']); hdr.pack(fill=tk.X)
        tk.Label(hdr, text="CONFIGURACIÓN DE LA FALLA", font=font(12, "bold"), bg=C['red'], fg='white', pady=12).pack()
        main = tk.Frame(tab, bg=C['panel']); main.pack(padx=40, pady=40, fill=tk.BOTH, expand=True)
        f1 = tk.Frame(main, bg=C['panel']); f1.pack(fill=tk.X, pady=(0, 24))
        tk.Label(f1, text="Barra fallada:", font=font(11, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(0, 12))
        self.combo_barra_falla = ttk.Combobox(f1, values=[str(i+1) for i in range(n)], state="readonly", width=8, font=font(10)); self.combo_barra_falla.pack(side=tk.LEFT); self.combo_barra_falla.set("1")
        f2 = tk.Frame(main, bg=C['panel']); f2.pack(fill=tk.X, pady=(0, 24))
        tk.Label(f2, text="Tipo de falla:", font=font(11, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(0, 12))
        self.combo_tipo_falla = ttk.Combobox(f2, values=["Monofásica (SLG)", "Línea-Línea (LL)", "Doble Línea-Tierra (LLG)", "Trifásica (LLL)"], state="readonly", width=26, font=font(10)); self.combo_tipo_falla.pack(side=tk.LEFT); self.combo_tipo_falla.set("Monofásica (SLG)")
        f3 = tk.Frame(main, bg=C['panel']); f3.pack(fill=tk.X, pady=(0, 24))
        tk.Label(f3, text="Z_F (fase):", font=font(11, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(0, 12))
        self.entry_ZF = ttk.Entry(f3, width=10, justify='center', font=font(10)); self.entry_ZF.pack(side=tk.LEFT, padx=(0, 4)); self.entry_ZF.insert(0, "0.0")
        tk.Label(f3, text="p.u.", font=font(9, "bold"), bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT, padx=(0, 30))
        tk.Label(f3, text="Z_g (tierra):", font=font(11, "bold"), bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(0, 12))
        self.entry_Zg = ttk.Entry(f3, width=10, justify='center', font=font(10)); self.entry_Zg.pack(side=tk.LEFT, padx=(0, 4)); self.entry_Zg.insert(0, "0.0")
        tk.Label(f3, text="p.u.", font=font(9, "bold"), bg=C['panel'], fg=C['accent']).pack(side=tk.LEFT)

    def _on_tipo_changed(self, idx):
        if self._updating: return
        try: self._dibujar_preview(); self._update_sidebar()
        except Exception: pass

    def _dibujar_preview(self):
        canvas = getattr(self, 'preview_canvas', None)
        if canvas is None: return
        try: n = int(self.entry_n.get())
        except Exception: return
        if n < 2 or len(self.tipo_barras) != n: return
        canvas.delete("all"); W = canvas.winfo_width(); H = canvas.winfo_height()
        if W < 50 or H < 50: W, H = 520, 520
        C = self.C
        canvas.create_rectangle(0, 0, W, H, fill='#FFFFFF', outline='')
        canvas.create_text(W / 2, 26, text="Vista previa — Sistema de fallas", font=font(13, "bold"), fill=C['primary'])
        canvas.create_line(W / 2 - 100, 42, W / 2 + 100, 42, fill=C['red'], width=2)
        tipos = [3 if "Slack" in cb.get() else (2 if "PV" in cb.get() else 1) for cb in self.tipo_barras]
        barra_falla = 0
        try:
            if self.combo_barra_falla: barra_falla = int(self.combo_barra_falla.get()) - 1
        except Exception: barra_falla = 0
        lineas = []
        for row in self.entries_lineas:
            a = self._parse_int(row[0].get(), 0); b = self._parse_int(row[1].get(), 0)
            if a > 0 and b > 0 and a <= n and b <= n and a != b: lineas.append((a - 1, b - 1))
        trafos = []
        for row in self.entries_trafos:
            a = self._parse_int(row[0].get(), 0); b = self._parse_int(row[1].get(), 0)
            if a > 0 and b > 0 and a <= n and b <= n and a != b: trafos.append((a - 1, b - 1))
        if not lineas and not trafos:
            canvas.create_text(W / 2, H / 2, text="Sin líneas ni trafos definidos\n\nComplete las pestañas para visualizar.", font=font(11), fill=C['text_ll'], justify=tk.CENTER)
            return
        cx, cy = W / 2, H / 2 + 20; R = min(W, H) * 0.30
        pos = [(cx + R * math.cos(2 * math.pi * i / n - math.pi / 2), cy + R * math.sin(2 * math.pi * i / n - math.pi / 2)) for i in range(n)]
        for (a, b) in lineas:
            x1, y1 = pos[a]; x2, y2 = pos[b]
            canvas.create_line(x1, y1, x2, y2, width=2.5, fill=C['gray_dark'], capstyle=tk.ROUND)
        for (a, b) in trafos:
            x1, y1 = pos[a]; x2, y2 = pos[b]
            canvas.create_line(x1, y1, x2, y2, width=3, fill=C['primary'], capstyle=tk.ROUND)
            mx, my = (x1+x2)/2, (y1+y2)/2
            canvas.create_oval(mx-6, my-6, mx+6, my+6, outline=C['primary'], width=2)
            canvas.create_oval(mx+6, my-6, mx+18, my+6, outline=C['primary'], width=2)
        color_map = {1: C['pq'], 2: C['pv'], 3: C['slack']}; ring_map = {1: '#FFEDD5', 2: '#EDE9FE', 3: '#D1FAE5'}; name_map = {1: 'PQ', 2: 'PV', 3: 'SLK'}
        for i, (x, y) in enumerate(pos):
            tipo = tipos[i]; col = color_map[tipo]; ring = ring_map[tipo]
            canvas.create_line(x - 24, y, x + 24, y, width=10, fill=C['gray_dark'], capstyle=tk.ROUND)
            canvas.create_oval(x - 17, y - 17, x + 17, y + 17, fill=ring, outline='')
            canvas.create_oval(x - 13, y - 13, x + 13, y + 13, fill=col, outline='white', width=2)
            if i == barra_falla: canvas.create_oval(x - 20, y - 20, x + 20, y + 20, outline=C['red'], width=3)
            canvas.create_text(x, y, text=str(i + 1), fill='white', font=font(10, "bold"))
            canvas.create_text(x, y + 30, text=name_map[tipo], font=font(8, "bold"), fill=col)
            if i == barra_falla:
                pts = [x - 6, y - 45, x + 2, y - 30, x - 3, y - 30, x + 1, y - 20, x - 8, y - 33, x - 3, y - 33, x - 6, y - 45]
                canvas.create_polygon(pts, fill=C['red'], outline='', width=2)

    # =================================================================
    # LÓGICA DE CÁLCULO
    # =================================================================
    def calcular_falla(self):
        try:
            self._actualizar_bases(); n = int(self.entry_n.get())
            tipos = []; V_pref = []; ang_pref = []
            for i in range(n):
                ts = self.tipo_barras[i].get(); tipos.append(3 if "Slack" in ts else (2 if "PV" in ts else 1))
                V_pref.append(self._parse_float(self.V_pref_widgets[i].get(), 1.0))
                ang_pref.append(math.radians(self._parse_float(self.ang_pref_widgets[i].get(), 0.0)))

            lineas = []
            for row in self.entries_lineas:
                try:
                    a = self._parse_int(row[0].get(), 0); b = self._parse_int(row[1].get(), 0)
                    R1 = self._parse_float(row[2].get(), 0.0); X1 = self._parse_float(row[3].get(), 0.0); B1_2 = self._parse_float(row[4].get(), 0.0)
                    R0 = self._parse_float(row[5].get(), 0.0); X0 = self._parse_float(row[6].get(), 0.0)
                    if a > 0 and b > 0 and a <= n and b <= n and a != b:
                        Z1 = complex(R1, X1)
                        Z0 = complex(R0, X0) if (R0 != 0 or X0 != 0) else 3 * Z1
                        lineas.append({'a': a-1, 'b': b-1, 'Z1': Z1, 'Z2': Z1, 'Z0': Z0, 'B1_2': B1_2})
                except Exception: continue

            trafos = []
            for row in self.entries_trafos:
                try:
                    a = self._parse_int(row[0].get(), 0); b = self._parse_int(row[1].get(), 0)
                    R = self._parse_float(row[2].get(), 0.0); X = self._parse_float(row[3].get(), 0.0)
                    config = row[4].get()
                    if a > 0 and b > 0 and a <= n and b <= n and a != b:
                        Z = complex(R, X)
                        trafos.append({'a': a-1, 'b': b-1, 'Z1': Z, 'Z2': Z, 'Z0': Z, 'config': config})
                except Exception: continue

            inductores = []
            for row in self.entries_inductores:
                try:
                    bus = self._parse_int(row[0].get(), 0) - 1; Q = self._parse_float(row[1].get(), 0.0)
                    if 0 <= bus < n and Q != 0: inductores.append({'bus': bus, 'Q': Q})
                except Exception: continue

            capacitores = []
            for row in self.entries_capacitores:
                try:
                    bus = self._parse_int(row[0].get(), 0) - 1; Q = self._parse_float(row[1].get(), 0.0)
                    if 0 <= bus < n and Q != 0: capacitores.append({'bus': bus, 'Q': Q})
                except Exception: continue

            cargas = [] 
            for row in self.entries_cargas:
                try:
                    bus = self._parse_int(row[0].get(), 0) - 1
                    P = self._parse_float(row[1].get(), 0.0)
                    Q = self._parse_float(row[2].get(), 0.0)
                    if 0 <= bus < n and (P != 0 or Q != 0): cargas.append({'bus': bus, 'P': P, 'Q': Q})
                except Exception: continue

            generadores = []
            for g in self.gen_widgets:
                try:
                    bus = self._parse_int(g[0].get(), 0) - 1
                    R1 = self._parse_float(g[1].get(), 0.0); X1 = self._parse_float(g[2].get(), 0.0)
                    R2 = self._parse_float(g[3].get(), 0.0); X2 = self._parse_float(g[4].get(), 0.0)
                    R0 = self._parse_float(g[5].get(), 0.0); X0 = self._parse_float(g[6].get(), 0.0)
                    if 0 <= bus < n and (abs(X1)>1e-9 or abs(X0)>1e-9):
                        generadores.append({'bus': bus, 'Z1': complex(R1,X1), 'Z2': complex(R2,X2), 'Z0': complex(R0,X0)})
                except Exception: continue

            barra_falla = int(self.combo_barra_falla.get()) - 1
            tipo_str = self.combo_tipo_falla.get()
            ZF = self._parse_float(self.entry_ZF.get(), 0.0) + 0j # Solo asumo resistencia para ZF
            Zg = self._parse_float(self.entry_Zg.get(), 0.0) + 0j
            
            tipo_falla = 'SLG'
            if tipo_str.startswith("Monofásica"): tipo_falla = 'SLG'
            elif tipo_str.startswith("Línea-Línea") and 'LLG' not in tipo_str: tipo_falla = 'LL'
            elif "Doble" in tipo_str or "LLG" in tipo_str: tipo_falla = 'LLG'
            else: tipo_falla = 'LLL'

            Ybus0 = self._construir_ybus(n, lineas, trafos, inductores, capacitores, generadores, cargas, seq=0)
            Ybus1 = self._construir_ybus(n, lineas, trafos, inductores, capacitores, generadores, cargas, seq=1)
            Ybus2 = self._construir_ybus(n, lineas, trafos, inductores, capacitores, generadores, cargas, seq=2)

            try: Zbus0 = np.linalg.inv(Ybus0)
            except np.linalg.LinAlgError: 
                messagebox.showerror("Error", "Ybus secuencia 0 singular. Revise conexiones de trafos.")
                return
            try: Zbus1 = np.linalg.inv(Ybus1)
            except np.linalg.LinAlgError: return
            try: Zbus2 = np.linalg.inv(Ybus2)
            except np.linalg.LinAlgError: return

            V_pref_c = V_pref[barra_falla] * complex(math.cos(ang_pref[barra_falla]), math.sin(ang_pref[barra_falla]))
            Zth0 = Zbus0[barra_falla][barra_falla]; Zth1 = Zbus1[barra_falla][barra_falla]; Zth2 = Zbus2[barra_falla][barra_falla]

            if tipo_falla == 'SLG': res = self._resolver_slg(V_pref_c, Zth0, Zth1, Zth2, ZF, Zg)
            elif tipo_falla == 'LLL': res = self._resolver_lll(V_pref_c, Zth1, ZF)
            elif tipo_falla == 'LL': res = self._resolver_ll(V_pref_c, Zth1, Zth2, ZF)
            elif tipo_falla == 'LLG': res = self._resolver_llg(V_pref_c, Zth0, Zth1, Zth2, ZF, Zg)

            I0 = res['I0']; I1 = res['I1']; I2 = res['I2']

            V0_bus = [-Zbus0[i][barra_falla] * I0 for i in range(n)]
            V1_bus = []
            for i in range(n):
                Vp_i = V_pref[i] * complex(math.cos(ang_pref[i]), math.sin(ang_pref[i]))
                V1_bus.append(Vp_i - Zbus1[i][barra_falla] * I1)
            V2_bus = [-Zbus2[i][barra_falla] * I2 for i in range(n)]

            a_op = complex(-0.5, math.sqrt(3)/2); a2 = a_op ** 2
            Ia = I0 + I1 + I2; Ib = I0 + a2 * I1 + a_op * I2; Ic = I0 + a_op * I1 + a2 * I2

            V0F = V0_bus[barra_falla]; V1F = V1_bus[barra_falla]; V2F = V2_bus[barra_falla]
            Va_F = V0F + V1F + V2F; Vb_F = V0F + a2 * V1F + a_op * V2F; Vc_F = V0F + a_op * V1F + a2 * V2F

            self.resultados = {
                'n': n, 'tipos': tipos, 'V_pref': V_pref, 'ang_pref': ang_pref,
                'lineas': lineas, 'trafos': trafos, 'generadores': generadores,
                'barra_falla': barra_falla, 'tipo_falla': tipo_falla, 'ZF': ZF, 'Zg': Zg,
                'Ybus0': Ybus0, 'Ybus1': Ybus1, 'Ybus2': Ybus2,
                'Zbus0': Zbus0, 'Zbus1': Zbus1, 'Zbus2': Zbus2,
                'Zth0': Zth0, 'Zth1': Zth1, 'Zth2': Zth2, 'V_pref_c': V_pref_c,
                'I0': I0, 'I1': I1, 'I2': I2, 'Ia': Ia, 'Ib': Ib, 'Ic': Ic,
                'V0_bus': V0_bus, 'V1_bus': V1_bus, 'V2_bus': V2_bus,
                'Va_F': Va_F, 'Vb_F': Vb_F, 'Vc_F': Vc_F, 'formula': res,
                'BASE_MVA': self.BASE_MVA, 'V_BASE_KV': self.V_BASE_KV,
            }

            self._set_status("Falla calculada correctamente.", 'ok')
            self._update_sidebar(); self.mostrar_resultados()
        except Exception as e:
            import traceback; traceback.print_exc()
            messagebox.showerror("Error", f"Error en el cálculo:\n{e}")

    def _construir_ybus(self, n, lineas, trafos, inductores, capacitores, generadores, cargas, seq):
        Y = np.zeros((n, n), dtype=complex)
        Sbase = self.BASE_MVA

        # 1. Líneas
        for L in lineas:
            Z = L['Z1'] if seq in (1, 2) else L['Z0']
            if abs(Z) < 1e-9: continue
            y = 1.0 / Z
            a, b = L['a'], L['b']
            Y[a][a] += y; Y[b][b] += y; Y[a][b] -= y; Y[b][a] -= y
            if seq in (1, 2):
                b_shunt = L.get('B1_2', 0.0)
                Y[a][a] += 1j * b_shunt; Y[b][b] += 1j * b_shunt

        # 2. Transformadores
        for T in trafos:
            Z = T['Z1'] if seq in (1, 2) else T['Z0']
            if abs(Z) < 1e-9: continue
            y = 1.0 / Z
            a, b = T['a'], T['b']; config = T['config']
            if seq in (1, 2):
                Y[a][a] += y; Y[b][b] += y; Y[a][b] -= y; Y[b][a] -= y
            else:
                if "YNyn" in config: Y[a][a] += y; Y[b][b] += y; Y[a][b] -= y; Y[b][a] -= y
                elif "YNd" in config: Y[a][a] += y
                elif "Dyn" in config: Y[b][b] += y

        # 3. Inductores
        for ind in inductores:
            i = ind['bus']; Q = ind['Q']
            Y[i][i] += -1j * (Q / Sbase)

        # 4. Capacitores
        if seq in (1, 2):
            for cap in capacitores:
                i = cap['bus']; Q = cap['Q']
                Y[i][i] += 1j * (Q / Sbase)

        # 5. Cargas de impedancia constante
        if seq in (1, 2):
            for load in cargas:
                i = load['bus']
                # Y_load = (P - jQ) / Sbase (en p.u. a tensión nominal 1.0)
                Y_load = complex(load['P'], -load['Q']) / Sbase
                Y[i][i] += Y_load

        # 6. Generadores
        for g in generadores:
            Z = g['Z1'] if seq in (1, 2) else g['Z0']
            if abs(Z) < 1e-9: continue
            y = 1.0 / Z; i = g['bus']
            Y[i][i] += y

        return Y
    def _resolver_slg(self, V_pref, Z0, Z1, Z2, ZF, Zg):
        denom = Z0 + Z1 + Z2 + 3 * ZF + 3 * Zg
        I1 = V_pref / denom if abs(denom) > 1e-12 else 0j
        return {'I0': I1, 'I1': I1, 'I2': I1, 'denominador': denom, 'tipo': 'SLG'}

    def _resolver_lll(self, V_pref, Z1, ZF):
        denom = Z1 + ZF
        I1 = V_pref / denom if abs(denom) > 1e-12 else 0j
        return {'I0': 0j, 'I1': I1, 'I2': 0j, 'denominador': denom, 'tipo': 'LLL'}

    def _resolver_ll(self, V_pref, Z1, Z2, ZF):
        denom = Z1 + Z2 + ZF
        I1 = V_pref / denom if abs(denom) > 1e-12 else 0j
        return {'I0': 0j, 'I1': I1, 'I2': -I1, 'denominador': denom, 'tipo': 'LL'}

    def _resolver_llg(self, V_pref, Z0, Z1, Z2, ZF, Zg):
        Za = Z0 + 3 * Zg + ZF; Zb = Z2 + ZF
        denom_02 = Z0 + Z2 + 2 * ZF + 3 * Zg
        ZT = (Za * Zb) / (Za + Zb) if abs(Za + Zb) > 1e-12 else 0j
        denom1 = Z1 + ZF + ZT
        I1 = V_pref / denom1 if abs(denom1) > 1e-12 else 0j
        if abs(denom_02) > 1e-12:
            I0 = -I1 * (Z2 + ZF) / denom_02
            I2 = -I1 * (Z0 + 3 * Zg + ZF) / denom_02
        else: I0 = 0j; I2 = 0j
        return {'I0': I0, 'I1': I1, 'I2': I2, 'ZT': ZT, 'Za': Za, 'Zb': Zb, 'denominador': denom1, 'denominador_02': denom_02, 'tipo': 'LLG'}

    def _calc_aportes_generadores(self):
        r = self.resultados; n = r['n']; k = r['barra_falla']
        I0, I1, I2 = r['I0'], r['I1'], r['I2']
        Zbus0, Zbus1, Zbus2 = r['Zbus0'], r['Zbus1'], r['Zbus2']
        a_op = complex(-0.5, math.sqrt(3)/2); a2 = a_op ** 2
        aportes = []
        for g in r['generadores']:
            i = g['bus']
            Z0g, Z1g, Z2g = g['Z0'], g['Z1'], g['Z2']
            dV0 = Zbus0[i][k] * I0; dV1 = Zbus1[i][k] * I1; dV2 = Zbus2[i][k] * I2
            Ig0 = dV0 / Z0g if abs(Z0g) > 1e-9 else 0j
            Ig1 = dV1 / Z1g if abs(Z1g) > 1e-9 else 0j
            Ig2 = dV2 / Z2g if abs(Z2g) > 1e-9 else 0j
            Iga = Ig0 + Ig1 + Ig2; Igb = Ig0 + a2 * Ig1 + a_op * Ig2; Igc = Ig0 + a_op * Ig1 + a2 * Ig2
            aportes.append({'bus': i + 1, 'Z0': Z0g, 'Z1': Z1g, 'Z2': Z2g, 'I0': Ig0, 'I1': Ig1, 'I2': Ig2, 'Ia': Iga, 'Ib': Igb, 'Ic': Igc})
        return aportes

    def _fmt_complex(self, z, decimals=4):
        if abs(z.imag) < 1e-12: return f"{z.real:.{decimals}f}"
        if abs(z.real) < 1e-12: return f"{z.imag:.{decimals}f}j"
        sign = "+" if z.imag >= 0 else "-"
        return f"{z.real:.{decimals}f} {sign} {abs(z.imag):.{decimals}f}j"

    # =================================================================
    # VENTANA DE RESULTADOS Y EXPORTACIÓN
    # =================================================================
    def mostrar_resultados(self):
        if not self.resultados: messagebox.showinfo("Info", "Calcula primero la falla."); return
        r = self.resultados; C = self.C; r['aportes'] = self._calc_aportes_generadores()
        win = tk.Toplevel(self.root); win.title(f"Resultados — Falla {r['tipo_falla']} en Barra {r['barra_falla']+1}")
        win.geometry("1200x820"); win.minsize(1050, 680); win.configure(bg=C['bg'])
        hdr = tk.Frame(win, bg=C['primary_d'], height=70); hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['red'], height=4).pack(fill=tk.X, side=tk.TOP)
        tk.Label(hdr, text=f"  RESULTADOS — FALLA {r['tipo_falla']} EN BARRA {r['barra_falla']+1}", font=font(13, "bold"), bg=C['primary_d'], fg='white').pack(side=tk.LEFT, padx=20, pady=22)
        tk.Button(hdr, text="✕", command=win.destroy, bg=C['danger'], fg='white', font=font(11, "bold"), relief=tk.FLAT, padx=14, pady=6, cursor='hand2', borderwidth=0).pack(side=tk.RIGHT, padx=12, pady=22)
        style = ttk.Style(); style.configure("Res.TNotebook", background=C['bg'], borderwidth=0)
        style.configure("Res.TNotebook.Tab", background=C['head_bg'], foreground=C['text_l'], padding=[14, 8], font=font(10, "bold"))
        style.map("Res.TNotebook.Tab", background=[('selected', C['primary'])], foreground=[('selected', 'white')])
        nb = ttk.Notebook(win, style="Res.TNotebook"); nb.pack(fill='both', expand=True, padx=14, pady=10)
        tabs = {'Resumen': self._tab_resumen, 
                'Procedimiento': self._tab_procedimiento, 
                'Ybus / Zbus': self._tab_ybus_zbus,
                  'Fórmulas': self._tab_formulas, 
                  'Reconstrucción': self._tab_reconstruccion,
                    'Aportes gen.': self._tab_aportes, 
                    'C. Ramas': self._tab_corrientes_ramas,  
                    'S. Ramas': self._tab_potencias_ramas,  
                    'Falla en Línea': self._tab_falla_linea,
                    'Teoría': self._tab_teoria}
        for nombre, builder in tabs.items():
            frame = tk.Frame(nb, bg=C['bg']); nb.add(frame, text=nombre); builder(frame)
        foot = tk.Frame(win, bg=C['surface'], height=50); foot.pack(fill=tk.X, side=tk.BOTTOM)
        tk.Button(foot, text="Cerrar", command=win.destroy, bg=C['primary'], fg='white', font=font(10, "bold"), relief=tk.FLAT, padx=24, pady=8, cursor='hand2').pack(side=tk.RIGHT, padx=16, pady=8)
        tk.Button(foot, text="📊 Exportar Excel", command=self._exportar_excel, bg=C['success'], fg='white', font=font(10, "bold"), relief=tk.FLAT, padx=20, pady=8, cursor='hand2').pack(side=tk.RIGHT, padx=4, pady=8)

    def _tab_resumen(self, parent):
        r = self.resultados; C = self.C; BASE = r['BASE_MVA']; V_BASE = r['V_BASE_KV']; I_base_kA = BASE / (math.sqrt(3) * V_BASE)
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0); sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg']); inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw'); canvas.configure(yscrollcommand=sb.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas.find_all()[0], width=e.width))
        sb.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        info = tk.Frame(inner, bg=C['surface'], highlightthickness=1, highlightbackground=C['border']); info.pack(fill=tk.X, padx=20, pady=(16, 12))
        tipo_txt = {'SLG': "Monofásica (1F-Tierra)", 'LL': "Línea-Línea", 'LLG': "Doble Línea-Tierra", 'LLL': "Trifásica"}[r['tipo_falla']]
        tk.Label(info, text=f"Tipo: {tipo_txt}   ·   Barra: {r['barra_falla']+1}   ·   Z_F = {r['ZF']:.4f} p.u.   ·   Z_g = {r['Zg']:.4f} p.u.", font=font_m(10), bg=C['surface'], fg=C['text'], padx=14, pady=10, anchor='w').pack(fill=tk.X)
        kpis = tk.Frame(inner, bg=C['bg']); kpis.pack(fill=tk.X, padx=20, pady=(0, 14))
        def _card(t, v, s, col):
            card = tk.Frame(kpis, bg=C['panel'], highlightthickness=1, highlightbackground=C['border']); card.pack(side=tk.LEFT, padx=6, fill=tk.Y)
            tk.Frame(card, bg=col, width=5).pack(side=tk.LEFT, fill=tk.Y)
            body = tk.Frame(card, bg=C['panel']); body.pack(side=tk.LEFT, padx=14, pady=10)
            tk.Label(body, text=t.upper(), font=font(8, "bold"), bg=C['panel'], fg=C['text_l']).pack(anchor='w')
            tk.Label(body, text=v, font=font(12, "bold"), bg=C['panel'], fg=col).pack(anchor='w', pady=(4, 0))
            tk.Label(body, text=s, font=font(8, italic=True), bg=C['panel'], fg=C['text_l']).pack(anchor='w')
        Ia_kA = abs(r['Ia']) * I_base_kA
        _card("Ifalla (fase a)", f"{Ia_kA:.3f} kA", f"{abs(r['Ia']):.4f} p.u.", C['red'])
        _card("Z₁ Thévenin", f"{abs(r['Zth1']):.4f}", f"{self._fmt_complex(r['Zth1'])} p.u.", C['accent'])
        _card("Z₀ Thévenin", f"{abs(r['Zth0']):.4f}", f"{self._fmt_complex(r['Zth0'])} p.u.", C['info'])
        _card("Z₂ Thévenin", f"{abs(r['Zth2']):.4f}", f"{self._fmt_complex(r['Zth2'])} p.u.", C['primary'])
        _card("I_base", f"{I_base_kA:.3f} kA", f"S={BASE:.0f} MVA · V={V_BASE:.1f} kV", C['text_l'])
        self._tabla_resultados(inner, "CORRIENTES EN LA BARRA FALLADA", ["Cantidad", "Módulo (p.u.)", "Módulo (kA)", "Ángulo (°)"],
            [["I₀ (sec. cero)", f"{abs(r['I0']):.4f}", f"{abs(r['I0'])*I_base_kA:.4f}", f"{math.degrees(cmath.phase(r['I0'])):+.2f}"],
             ["I₁ (sec. positiva)", f"{abs(r['I1']):.4f}", f"{abs(r['I1'])*I_base_kA:.4f}", f"{math.degrees(cmath.phase(r['I1'])):+.2f}"],
             ["I₂ (sec. negativa)", f"{abs(r['I2']):.4f}", f"{abs(r['I2'])*I_base_kA:.4f}", f"{math.degrees(cmath.phase(r['I2'])):+.2f}"],
             ["Iₐ (fase a)", f"{abs(r['Ia']):.4f}", f"{abs(r['Ia'])*I_base_kA:.4f}", f"{math.degrees(cmath.phase(r['Ia'])):+.2f}"],
             ["I_b (fase b)", f"{abs(r['Ib']):.4f}", f"{abs(r['Ib'])*I_base_kA:.4f}", f"{math.degrees(cmath.phase(r['Ib'])):+.2f}"],
             ["I_c (fase c)", f"{abs(r['Ic']):.4f}", f"{abs(r['Ic'])*I_base_kA:.4f}", f"{math.degrees(cmath.phase(r['Ic'])):+.2f}"]], color_col=1)

    def _tab_procedimiento(self, parent):
        r = self.resultados; C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0); sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg']); inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw'); canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        def _titulo(txt):
            f = tk.Frame(inner, bg=C['primary_d']); f.pack(fill=tk.X, padx=20, pady=(14, 6))
            tk.Frame(f, bg=C['red'], width=5).pack(side=tk.LEFT, fill=tk.Y)
            tk.Label(f, text=txt, font=font(11, "bold"), bg=C['primary_d'], fg='white', pady=8).pack(side=tk.LEFT, padx=12)
        def _txt(s):
            f = tk.Frame(inner, bg=C['surface'], highlightthickness=1, highlightbackground=C['border']); f.pack(fill=tk.X, padx=20, pady=3)
            tk.Label(f, text=s, font=font_m(10), bg=C['surface'], fg=C['text'], justify='left', anchor='w', padx=14, pady=10, wraplength=1050).pack(fill=tk.X)
        def _formula(s):
            f = tk.Frame(inner, bg='#1e293b', highlightthickness=1, highlightbackground=C['accent']); f.pack(fill=tk.X, padx=20, pady=4)
            tk.Label(f, text=s, font=('Consolas', 10, 'bold'), bg='#1e293b', fg='#7dd3fc', justify='left', anchor='w', padx=14, pady=10).pack(fill=tk.X)
        _titulo("PASO 1 — DATOS DE ENTRADA")
        _txt(f"Barras: {r['n']}    Líneas: {len(r['lineas'])}    Trafos: {len(r['trafos'])}    Generadores: {len(r['generadores'])}")
        _titulo("PASO 2 — CONSTRUCCIÓN DE LAS 3 Ybus")
        _formula("Y_ii += 1/Z_k     Y_jj += 1/Z_k\nY_ij -= 1/Z_k     Y_ji -= 1/Z_k\nY_ii += jB/2 (en líneas)\nY_ii += Y_shunt")
        _titulo("PASO 3 — INVERSIÓN: Zbus = Ybus⁻¹")
        _formula(f"Z₀_th = {self._fmt_complex(r['Zth0'])} p.u.\nZ₁_th = {self._fmt_complex(r['Zth1'])} p.u.\nZ₂_th = {self._fmt_complex(r['Zth2'])} p.u.")
        _titulo("PASO 4 — CONEXIÓN DE SECUENCIAS")
        formulas = {'SLG': "Las 3 secuencias EN SERIE:\nI₀ = I₁ = I₂ = V₁⁰ / (Z₀ + Z₁ + Z₂ + 3Z_F + 3Z_g)",
            'LL':  "Sec. 1 y 2 en PARALELO, sec. 0 abierta:\nI₁ = V₁⁰ / (Z₁ + Z₂ + Z_F);  I₂ = -I₁;  I₀ = 0",
            'LLL': "Solo secuencia positiva:\nI₁ = V₁⁰ / (Z₁ + Z_F);  I₀ = I₂ = 0",
            'LLG': "Las 3 secuencias en PARALELO:\nZ_T = (Z₀ + 3Z_g + Z_F) || (Z₂ + Z_F)\nI₁ = V₁⁰ / (Z₁ + Z_F + Z_T)\nI₀ = -I₁·(Z₂+Z_F) / (Z₀+Z₂+2Z_F+3Z_g)\nI₂ = -I₁·(Z₀+3Z_g+Z_F) / (Z₀+Z₂+2Z_F+3Z_g)"}
        _formula(formulas[r['tipo_falla']])
        _titulo("PASO 5 — CORRIENTES DE SECUENCIA")
        _formula(f"I₀ = {self._fmt_complex(r['I0'])} p.u.\nI₁ = {self._fmt_complex(r['I1'])} p.u.\nI₂ = {self._fmt_complex(r['I2'])} p.u.")

    def _tab_ybus_zbus(self, parent):
        r = self.resultados; C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0); sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg']); inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw'); canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        for titulo, M in [("Ybus — SECUENCIA 0", r['Ybus0']), ("Ybus — SECUENCIA 1", r['Ybus1']), ("Ybus — SECUENCIA 2", r['Ybus2']),
                          ("Zbus — SECUENCIA 0", r['Zbus0']), ("Zbus — SECUENCIA 1", r['Zbus1']), ("Zbus — SECUENCIA 2", r['Zbus2'])]:
            self._mostrar_matriz(inner, titulo, M)

    def _mostrar_matriz(self, parent, titulo, M):
        C = self.C; n = len(M)
        f = tk.Frame(parent, bg=C['surface'], highlightthickness=1, highlightbackground=C['border']); f.pack(fill=tk.X, padx=20, pady=(14, 6))
        tk.Frame(f, bg=C['red'], width=5).pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(f, text=titulo, font=font(10, "bold"), bg=C['surface'], fg=C['text'], pady=6).pack(side=tk.LEFT, padx=12)
        tb = tk.Frame(parent, bg=C['border']); tb.pack(fill=tk.X, padx=20, pady=(0, 12))
        head = tk.Frame(tb, bg=C['primary_d']); head.pack(fill=tk.X)
        tk.Label(head, text="i\\j", font=font(9, "bold"), width=6, bg=C['primary_d'], fg='white', pady=6).pack(side=tk.LEFT)
        for j in range(n): tk.Label(head, text=f"j={j+1}", font=font(9, "bold"), width=22, bg=C['primary_d'], fg='white', pady=6).pack(side=tk.LEFT)
        for i in range(n):
            row = tk.Frame(tb, bg=C['border']); row.pack(fill=tk.X, pady=1)
            tk.Label(row, text=f"i={i+1}", font=font(9, "bold"), width=6, bg=C['row_odd'], fg=C['text'], pady=5).pack(side=tk.LEFT)
            for j in range(n):
                bg_c = C['row_even'] if (i + j) % 2 == 0 else C['row_odd']
                val = M[i][j]; txt = self._fmt_complex(val, 4) if isinstance(val, complex) else f"{val:.4f}"
                tk.Label(row, text=txt, font=('Consolas', 9), width=22, bg=bg_c, fg=C['text'], pady=5).pack(side=tk.LEFT)

    def _tab_formulas(self, parent):
        r = self.resultados; C = self.C; f = r['formula']
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0); sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg']); inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw'); canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        def _blk(t, contenido):
            tk.Label(inner, text=t, font=font(11, "bold"), bg=C['bg'], fg=C['primary'], anchor='w').pack(fill=tk.X, padx=20, pady=(14, 4))
            fr = tk.Frame(inner, bg='#1e293b', highlightthickness=1, highlightbackground=C['accent']); fr.pack(fill=tk.X, padx=20, pady=2)
            tk.Label(fr, text=contenido, font=('Consolas', 10), bg='#1e293b', fg='#7dd3fc', justify='left', anchor='w', padx=14, pady=10).pack(fill=tk.X)
        _blk("FÓRMULA GENERAL", f"Z₀_th = {self._fmt_complex(r['Zth0'])} p.u.\nZ₁_th = {self._fmt_complex(r['Zth1'])} p.u.\nZ₂_th = {self._fmt_complex(r['Zth2'])} p.u.\nV₁⁰  = {self._fmt_complex(r['V_pref_c'])} p.u.\nZ_F  = {r['ZF']:.4f} p.u.;  Z_g = {r['Zg']:.4f} p.u.")
        if r['tipo_falla'] == 'SLG': _blk("SUSTITUCIÓN (SLG)", f"I₁ = V₁⁰ / (Z₀+Z₁+Z₂+3Z_F+3Z_g) = {self._fmt_complex(r['V_pref_c'])} / {self._fmt_complex(f['denominador'])} = {self._fmt_complex(f['I1'])} p.u.")
        elif r['tipo_falla'] == 'LL': _blk("SUSTITUCIÓN (LL)", f"I₁ = V₁⁰ / (Z₁+Z₂+Z_F) = {self._fmt_complex(f['I1'])} p.u.\nI₂ = -I₁ = {self._fmt_complex(f['I2'])} p.u.;  I₀ = 0")
        elif r['tipo_falla'] == 'LLL': _blk("SUSTITUCIÓN (LLL)", f"I₁ = V₁⁰ / (Z₁+Z_F) = {self._fmt_complex(f['I1'])} p.u.;  I₀ = I₂ = 0")
        else: _blk("SUSTITUCIÓN (LLG)", f"Z_T = {self._fmt_complex(f['Za'])} || {self._fmt_complex(f['Zb'])} = {self._fmt_complex(f['ZT'])} p.u.\nI₁ = {self._fmt_complex(f['I1'])} p.u.\nI₀ = {self._fmt_complex(f['I0'])} p.u.\nI₂ = {self._fmt_complex(f['I2'])} p.u.")

    def _tab_reconstruccion(self, parent):
        r = self.resultados; C = self.C; BASE = r['BASE_MVA']; V_BASE = r['V_BASE_KV']; I_base = BASE / (math.sqrt(3) * V_BASE)
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0); sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg']); inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw'); canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        self._tabla_resultados(inner, "CORRIENTES abc EN LA BARRA FALLADA", ["Fase", "Módulo (p.u.)", "Módulo (kA)", "Ángulo (°)"],
            [["Iₐ", f"{abs(r['Ia']):.4f}", f"{abs(r['Ia'])*I_base:.4f}", f"{math.degrees(cmath.phase(r['Ia'])):+.2f}"],
             ["I_b", f"{abs(r['Ib']):.4f}", f"{abs(r['Ib'])*I_base:.4f}", f"{math.degrees(cmath.phase(r['Ib'])):+.2f}"],
             ["I_c", f"{abs(r['Ic']):.4f}", f"{abs(r['Ic'])*I_base:.4f}", f"{math.degrees(cmath.phase(r['Ic'])):+.2f}"]], color_col=1)
        n = r['n']; k = r['barra_falla']; a_op = complex(-0.5, math.sqrt(3)/2); a2 = a_op**2
        filas_seq, filas_abc = [], []
        for i in range(n):
            V0 = r['V0_bus'][i]; V1 = r['V1_bus'][i]; V2 = r['V2_bus'][i]
            Va = V0 + V1 + V2; Vb = V0 + a2*V1 + a_op*V2; Vc = V0 + a_op*V1 + a2*V2
            marca = " ← falla" if i == k else ""
            filas_seq.append([f"Barra {i+1}{marca}", f"{abs(V0):.4f}", f"{abs(V1):.4f}", f"{abs(V2):.4f}"])
            filas_abc.append([f"Barra {i+1}{marca}", f"{abs(Va):.4f}", f"{abs(Vb):.4f}", f"{abs(Vc):.4f}", f"{abs(Va)*V_BASE:.3f}", f"{abs(Vb)*V_BASE:.3f}", f"{abs(Vc)*V_BASE:.3f}"])
        self._tabla_resultados(inner, "TENSIONES DE SECUENCIA POR BARRA (p.u.)", ["Barra", "|V₀|", "|V₁|", "|V₂|"], filas_seq, ancho=14)
        self._tabla_resultados(inner, "TENSIONES abc POR BARRA", ["Barra", "|Va| p.u.", "|Vb| p.u.", "|Vc| p.u.", "|Va| kV", "|Vb| kV", "|Vc| kV"], filas_abc, ancho=13)

    def _tab_aportes(self, parent):
        r = self.resultados; C = self.C; BASE = r['BASE_MVA']; V_BASE = r['V_BASE_KV']; I_base = BASE / (math.sqrt(3) * V_BASE)
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0); sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg']); inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw'); canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        aportes = r.get('aportes', [])
        if not aportes: tk.Label(inner, text="No hay generadores definidos con Z > 0.", font=font(11, italic=True), bg=C['bg'], fg=C['text_l'], pady=40).pack(); return
        filas_seq = [[f"Bus {a['bus']}", self._fmt_complex(a['Z1']), f"{abs(a['I0']):.4f}", f"{abs(a['I1']):.4f}", f"{abs(a['I2']):.4f}"] for a in aportes]
        self._tabla_resultados(inner, "CORRIENTES DE SECUENCIA POR GENERADOR (p.u.)", ["Generador", "Z₁ (p.u.)", "|I₀|", "|I₁|", "|I₂|"], filas_seq, color_col=3, ancho=16)
        filas_abc = [[f"Bus {a['bus']}", f"{abs(a['Ia']):.4f}", f"{abs(a['Ib']):.4f}", f"{abs(a['Ic']):.4f}", f"{abs(a['Ia'])*I_base:.3f}", f"{abs(a['Ib'])*I_base:.3f}", f"{abs(a['Ic'])*I_base:.3f}"] for a in aportes]
        self._tabla_resultados(inner, "APORTES abc POR GENERADOR", ["Generador", "|Ia| p.u.", "|Ib| p.u.", "|Ic| p.u.", "|Ia| kA", "|Ib| kA", "|Ic| kA"], filas_abc, color_col=4, ancho=13)

    # =================================================================
    # CORRIENTES POR RAMAS (SECUENCIAS 012)
    # =================================================================
    def _tab_corrientes_ramas(self, parent):
        r = self.resultados; C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))

        def fmt(z):
            return f"{z.real:+.3f} + j{z.imag:+.3f}"

        a = complex(-0.5, math.sqrt(3)/2); a2 = a**2
        def to_abc(I0, I1, I2):
            return (I0 + I1 + I2), (I0 + a2*I1 + a*I2), (I0 + a*I1 + a2*I2)

        # 1. LÍNEAS
        filas_lineas = []
        for L in r['lineas']:
            a_bus, b_bus = L['a'], L['b']
            y1 = 1.0 / L['Z1']; y2 = 1.0 / L['Z2']; y0 = 1.0 / L['Z0']
            b_shunt = L.get('B1_2', 0.0)
            I1 = (r['V1_bus'][a_bus] - r['V1_bus'][b_bus]) * y1 + r['V1_bus'][a_bus] * (1j*b_shunt)
            I2 = (r['V2_bus'][a_bus] - r['V2_bus'][b_bus]) * y2 + r['V2_bus'][a_bus] * (1j*b_shunt)
            I0 = (r['V0_bus'][a_bus] - r['V0_bus'][b_bus]) * y0
            Ia, Ib, Ic = to_abc(I0, I1, I2)
            filas_lineas.append([f"{a_bus+1}-{b_bus+1}", fmt(I0), fmt(I1), fmt(I2), fmt(Ia), fmt(Ib), fmt(Ic)])
        self._tabla_resultados(inner, "CORRIENTES EN LÍNEAS", ["Rama", "I_0", "I_1", "I_2", "Ia", "Ib", "Ic"], filas_lineas, ancho=14)

        # 2. TRANSFORMADORES
        filas_trafos = []
        for T in r['trafos']:
            a_bus, b_bus = T['a'], T['b']; config = T['config']
            y1 = 1.0 / T['Z1']; y2 = 1.0 / T['Z2']; y0 = 1.0 / T['Z0']
            I1 = (r['V1_bus'][a_bus] - r['V1_bus'][b_bus]) * y1
            I2 = (r['V2_bus'][a_bus] - r['V2_bus'][b_bus]) * y2
            if "YNyn" in config: I0 = (r['V0_bus'][a_bus] - r['V0_bus'][b_bus]) * y0
            elif "YNd" in config: I0 = r['V0_bus'][a_bus] * y0
            elif "Dyn" in config: I0 = -r['V0_bus'][b_bus] * y0
            else: I0 = 0j
            Ia, Ib, Ic = to_abc(I0, I1, I2)
            filas_trafos.append([f"{a_bus+1}-{b_bus+1} ({config.split()[0]})", fmt(I0), fmt(I1), fmt(I2), fmt(Ia), fmt(Ib), fmt(Ic)])
        self._tabla_resultados(inner, "CORRIENTES EN TRANSFORMADORES", ["Rama", "I_0", "I_1", "I_2", "Ia", "Ib", "Ic"], filas_trafos, ancho=14)

        # 3. SHUNTS (Reactores y Capacitores)
        filas_shunts = []
        for ind in r.get('inductores', []):
            i = ind['bus']; Q = ind['Q']; Y_sh = -1j * (Q / r['BASE_MVA'])
            I0 = r['V0_bus'][i] * Y_sh; I1 = r['V1_bus'][i] * Y_sh; I2 = r['V2_bus'][i] * Y_sh
            Ia, Ib, Ic = to_abc(I0, I1, I2)
            filas_shunts.append([f"Inductor N{i+1}", fmt(I0), fmt(I1), fmt(I2), fmt(Ia), fmt(Ib), fmt(Ic)])
        for cap in r.get('capacitores', []):
            i = cap['bus']; Q = cap['Q']; Y_sh = 1j * (Q / r['BASE_MVA'])
            I0 = r['V0_bus'][i] * Y_sh; I1 = r['V1_bus'][i] * Y_sh; I2 = r['V2_bus'][i] * Y_sh
            Ia, Ib, Ic = to_abc(I0, I1, I2)
            filas_shunts.append([f"Capacitor N{i+1}", fmt(I0), fmt(I1), fmt(I2), fmt(Ia), fmt(Ib), fmt(Ic)])
        if filas_shunts:
            self._tabla_resultados(inner, "CORRIENTES EN SHUNTS", ["Elemento", "I_0", "I_1", "I_2", "Ia", "Ib", "Ic"], filas_shunts, ancho=14)

        # 4. GENERADORES (Inyección al nodo)
        filas_gens = []
        for g in r['generadores']:
            i = g['bus']
            Z1, Z2, Z0 = g['Z1'], g['Z2'], g['Z0']
            Vpref_bus = r['V_pref'][i] * complex(math.cos(r['ang_pref'][i]), math.sin(r['ang_pref'][i]))
            # Corriente del generador = (Voltaje interno - Voltaje nodo) / Z
            I1 = (Vpref_bus - r['V1_bus'][i]) / Z1 if abs(Z1) > 1e-9 else 0j
            I2 = (0 - r['V2_bus'][i]) / Z2 if abs(Z2) > 1e-9 else 0j
            I0 = (0 - r['V0_bus'][i]) / Z0 if abs(Z0) > 1e-9 else 0j
            Ia, Ib, Ic = to_abc(I0, I1, I2)
            filas_gens.append([f"Gen en N{i+1}", fmt(I0), fmt(I1), fmt(I2), fmt(Ia), fmt(Ib), fmt(Ic)])
        if filas_gens:
            self._tabla_resultados(inner, "CORRIENTES INYECTADAS POR GENERADORES", ["Elemento", "I_0", "I_1", "I_2", "Ia", "Ib", "Ic"], filas_gens, ancho=14)

    # =================================================================
    # POTENCIA POR RAMAS (S = V * I*)
    # =================================================================
    def _tab_potencias_ramas(self, parent):
        r = self.resultados; C = self.C
        Sbase = r['BASE_MVA']
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))

        def fmt_pot(s_pu):
            s_mva = s_pu * Sbase
            return f"{s_mva.real:+.2f} + j{s_mva.imag:+.2f}"

        a = complex(-0.5, math.sqrt(3)/2); a2 = a**2
        def to_abc(V0, V1, V2):
            return (V0 + V1 + V2), (V0 + a2*V1 + a*V2), (V0 + a*V1 + a2*V2)

        # 1. LÍNEAS
        filas_lineas = []
        for L in r['lineas']:
            a_bus, b_bus = L['a'], L['b']
            y1 = 1.0 / L['Z1']; y2 = 1.0 / L['Z2']; y0 = 1.0 / L['Z0']
            b_shunt = L.get('B1_2', 0.0)
            I1 = (r['V1_bus'][a_bus] - r['V1_bus'][b_bus]) * y1 + r['V1_bus'][a_bus] * (1j*b_shunt)
            I2 = (r['V2_bus'][a_bus] - r['V2_bus'][b_bus]) * y2 + r['V2_bus'][a_bus] * (1j*b_shunt)
            I0 = (r['V0_bus'][a_bus] - r['V0_bus'][b_bus]) * y0
            Va, Vb, Vc = to_abc(r['V0_bus'][a_bus], r['V1_bus'][a_bus], r['V2_bus'][a_bus])
            Ia, Ib, Ic = to_abc(I0, I1, I2)
            Sa = Va * Ia.conjugate(); Sb = Vb * Ib.conjugate(); Sc = Vc * Ic.conjugate()
            filas_lineas.append([f"{a_bus+1}-{b_bus+1}", fmt_pot(Sa), fmt_pot(Sb), fmt_pot(Sc), fmt_pot(Sa+Sb+Sc)])
        self._tabla_resultados(inner, "POTENCIA EN LÍNEAS (MVA)", ["Rama", "S_a", "S_b", "S_c", "S_Total (3φ)"], filas_lineas, ancho=18)

        # 2. TRANSFORMADORES
        filas_trafos = []
        for T in r['trafos']:
            a_bus, b_bus = T['a'], T['b']; config = T['config']
            y1 = 1.0 / T['Z1']; y2 = 1.0 / T['Z2']; y0 = 1.0 / T['Z0']
            I1 = (r['V1_bus'][a_bus] - r['V1_bus'][b_bus]) * y1
            I2 = (r['V2_bus'][a_bus] - r['V2_bus'][b_bus]) * y2
            if "YNyn" in config: I0 = (r['V0_bus'][a_bus] - r['V0_bus'][b_bus]) * y0
            elif "YNd" in config: I0 = r['V0_bus'][a_bus] * y0
            elif "Dyn" in config: I0 = -r['V0_bus'][b_bus] * y0
            else: I0 = 0j
            Va, Vb, Vc = to_abc(r['V0_bus'][a_bus], r['V1_bus'][a_bus], r['V2_bus'][a_bus])
            Ia, Ib, Ic = to_abc(I0, I1, I2)
            Sa = Va * Ia.conjugate(); Sb = Vb * Ib.conjugate(); Sc = Vc * Ic.conjugate()
            filas_trafos.append([f"{a_bus+1}-{b_bus+1} ({config.split()[0]})", fmt_pot(Sa), fmt_pot(Sb), fmt_pot(Sc), fmt_pot(Sa+Sb+Sc)])
        self._tabla_resultados(inner, "POTENCIA EN TRANSFORMADORES (MVA)", ["Rama", "S_a", "S_b", "S_c", "S_Total (3φ)"], filas_trafos, ancho=18)

        # 3. GENERADORES
        filas_gens = []
        for g in r['generadores']:
            i = g['bus']
            Z1, Z2, Z0 = g['Z1'], g['Z2'], g['Z0']
            Vpref_bus = r['V_pref'][i] * complex(math.cos(r['ang_pref'][i]), math.sin(r['ang_pref'][i]))
            I1 = (Vpref_bus - r['V1_bus'][i]) / Z1 if abs(Z1) > 1e-9 else 0j
            I2 = (0 - r['V2_bus'][i]) / Z2 if abs(Z2) > 1e-9 else 0j
            I0 = (0 - r['V0_bus'][i]) / Z0 if abs(Z0) > 1e-9 else 0j
            Va, Vb, Vc = to_abc(r['V0_bus'][i], r['V1_bus'][i], r['V2_bus'][i])
            Ia, Ib, Ic = to_abc(I0, I1, I2)
            Sa = Va * Ia.conjugate(); Sb = Vb * Ib.conjugate(); Sc = Vc * Ic.conjugate()
            filas_gens.append([f"Gen en N{i+1}", fmt_pot(Sa), fmt_pot(Sb), fmt_pot(Sc), fmt_pot(Sa+Sb+Sc)])
        if filas_gens:
            self._tabla_resultados(inner, "POTENCIA ENTREGADA POR GENERADORES (MVA)", ["Elemento", "S_a", "S_b", "S_c", "S_Total (3φ)"], filas_gens, ancho=18)

    def _tab_teoria(self, parent):
        C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0); sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg']); inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw'); canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y'); canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))
        def _sec(titulo, texto):
            fr = tk.Frame(inner, bg=C['surface'], highlightthickness=1, highlightbackground=C['border']); fr.pack(fill=tk.X, padx=20, pady=8)
            tk.Frame(fr, bg=C['accent'], width=5).pack(side=tk.LEFT, fill=tk.Y)
            body = tk.Frame(fr, bg=C['surface']); body.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=14, pady=10)
            tk.Label(body, text=titulo, font=font(11, "bold"), bg=C['surface'], fg=C['primary'], anchor='w').pack(fill=tk.X)
            tk.Label(body, text=texto, font=font_m(10), bg=C['surface'], fg=C['text'], justify='left', anchor='w', wraplength=1050).pack(fill=tk.X, pady=(4, 0))
        _sec("COMPONENTES SIMÉTRICAS", "Descompone un sistema trifásico desequilibrado en tres sistemas equilibrados.\nSec. 0: 3 fases en fase.\nSec. 1: a-b-c.\nSec. 2: a-c-b.")
        _sec("FALLA MONOFÁSICA (SLG)", "I₀ = I₁ = I₂ = V₁⁰ / (Z₀ + Z₁ + Z₂ + 3Z_F + 3Z_g)")
        _sec("FALLA LÍNEA-LÍNEA (LL)", "I₁ = V₁⁰ / (Z₁ + Z₂ + Z_F);  I₂ = -I₁;  I₀ = 0")
        _sec("DOBLE LÍNEA-TIERRA (LLG)", "Z_T = (Z₀ + 3Z_g + Z_F) || (Z₂ + Z_F)\nI₁ = V₁⁰ / (Z₁ + Z_F + Z_T)")
        _sec("FALLA TRIFÁSICA (LLL)", "I₁ = V₁⁰ / (Z₁ + Z_F);  I₀ = I₂ = 0")

    def _tabla_resultados(self, parent, titulo, headers, rows, color_col=None, ancho=18):
        C = self.C
        f_titulo = tk.Frame(parent, bg=C['surface'], highlightthickness=1, highlightbackground=C['border']); f_titulo.pack(fill=tk.X, padx=20, pady=(10, 8))
        tk.Frame(f_titulo, bg=C['red'], width=5).pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(f_titulo, text=titulo, font=font(10, "bold"), bg=C['surface'], fg=C['text'], pady=8).pack(side=tk.LEFT, padx=12)
        tabla = tk.Frame(parent, bg=C['border']); tabla.pack(fill=tk.X, padx=20, pady=(0, 14))
        head = tk.Frame(tabla, bg=C['primary_d']); head.pack(fill=tk.X)
        for h in headers: tk.Label(head, text=h, font=font(9, "bold"), bg=C['primary_d'], fg='white', width=ancho, pady=8).pack(side=tk.LEFT)
        body = tk.Frame(tabla, bg=C['border']); body.pack(fill=tk.X)
        for i, fila in enumerate(rows):
            bg_r = C['row_even'] if i % 2 == 0 else C['row_odd']; f = tk.Frame(body, bg=C['border']); f.pack(fill=tk.X, pady=1)
            for c, val in enumerate(fila):
                fg_c = C['red'] if (color_col is not None and c == color_col) else C['text']
                tk.Label(f, text=str(val), font=font_m(9), bg=bg_r, fg=fg_c, width=ancho, pady=6).pack(side=tk.LEFT)

    def mostrar_procedimiento(self):
        if not self.resultados: messagebox.showinfo("Info", "Calcula primero la falla."); return
        self.mostrar_resultados()

    def mostrar_teoria(self):
        if not self.resultados: self._ventana_teoria_standalone(); return
        self.mostrar_resultados()

    def _ventana_teoria_standalone(self):
        C = self.C; win = tk.Toplevel(self.root); win.title("Teoría de Fallas — Vatia"); win.geometry("900x700"); win.configure(bg=C['bg'])
        hdr = tk.Frame(win, bg=C['primary_d'], height=64); hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="  TEORÍA DE FALLAS — COMPONENTES SIMÉTRICAS", font=font(14, "bold"), bg=C['primary_d'], fg='white').pack(side=tk.LEFT, padx=20, pady=18)
        tk.Button(hdr, text="✕", command=win.destroy, bg=C['danger'], fg='white', font=font(11, "bold"), relief=tk.FLAT, padx=14, pady=6, cursor='hand2').pack(side=tk.RIGHT, padx=12, pady=14)
        frame = tk.Frame(win, bg=C['bg']); frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        texto = "COMPONENTES SIMÉTRICAS\n\nSec. 0: 3 fases en fase\nSec. 1: a-b-c\nSec. 2: a-c-b\n\nFALLA MONOFÁSICA (SLG)\nI₀ = I₁ = I₂ = V₁⁰ / (Z₀ + Z₁ + Z₂ + 3Z_F + 3Z_g)\n\nFALLA LÍNEA-LÍNEA (LL)\nI₁ = V₁⁰ / (Z₁ + Z₂ + Z_F); I₂ = -I₁; I₀ = 0\n\nFALLA DOBLE LÍNEA-TIERRA (LLG)\nZ_T = (Z₀+3Z_g+Z_F) || (Z₂+Z_F)\nI₁ = V₁⁰ / (Z₁ + Z_F + Z_T)\n\nFALLA TRIFÁSICA (LLL)\nI₁ = V₁⁰ / (Z₁ + Z_F); I₀ = I₂ = 0"
        tk.Label(frame, text=texto, font=font_m(11), bg=C['bg'], fg=C['text'], justify='left', anchor='w').pack(fill=tk.BOTH, expand=True)

    def _nuevo_proyecto(self):
        if messagebox.askyesno("Nuevo proyecto", "¿Crear un nuevo proyecto?\n\nSe perderán los datos actuales."):
            self.generar_tablas(); self._set_status("Nuevo proyecto creado.", 'info')

    def _guardar_proyecto(self):
        path = filedialog.asksaveasfilename(defaultextension=".vatia_fallas.json", filetypes=[("Proyecto Fallas (VATIA)", "*.vatia_fallas.json"), ("JSON", "*.json")])
        if not path: return
        try:
            n = int(self.entry_n.get())
            data = {
                'version': 1, 'modo': 'fallas', 'fecha': datetime.now().isoformat(), 'n': n,
                'base_mva': self.entry_base_mva.get(), 'base_kv': self.entry_base_kv.get(),
                'barras': [{'tipo': self.tipo_barras[i].get(), 'V': self.V_pref_widgets[i].get(), 'ang': self.ang_pref_widgets[i].get()} for i in range(n)],
                'lineas': [[e.get() for e in row] for row in self.entries_lineas],
                'trafos': [[e.get() for e in row] for row in self.entries_trafos],
                'inductores': [[e.get() for e in row] for row in self.entries_inductores],
                'capacitores': [[e.get() for e in row] for row in self.entries_capacitores],
                'cargas': [[e.get() for e in row] for row in self.entries_cargas],
                'generadores': [[e.get() for e in row] for row in self.gen_widgets],
                'falla': {'barra': self.combo_barra_falla.get(), 'tipo': self.combo_tipo_falla.get(), 'ZF': self.entry_ZF.get(), 'Zg': self.entry_Zg.get()},
            }
            with open(path, 'w', encoding='utf-8') as f: json.dump(data, f, indent=2, ensure_ascii=False)
            self._set_status(f"Guardado: {os.path.basename(path)}", 'ok'); messagebox.showinfo("Guardado", f"Proyecto guardado en:\n{path}")
        except Exception as e: messagebox.showerror("Error al guardar", str(e))

    def _abrir_proyecto(self):
        path = filedialog.askopenfilename(filetypes=[("Proyecto Fallas (VATIA)", "*.vatia_fallas.json"), ("JSON", "*.json")])
        if not path: return
        try:
            with open(path, encoding='utf-8') as f: data = json.load(f)
            if 'base_mva' in data: self.entry_base_mva.delete(0, tk.END); self.entry_base_mva.insert(0, data['base_mva'])
            if 'base_kv' in data: self.entry_base_kv.delete(0, tk.END); self.entry_base_kv.insert(0, data['base_kv'])
            self._actualizar_bases()
            self.entry_n.delete(0, tk.END); self.entry_n.insert(0, str(data['n']))
            self.entry_n_lineas.delete(0, tk.END); self.entry_n_lineas.insert(0, str(len(data.get('lineas', []))))
            self.entry_n_trafos.delete(0, tk.END); self.entry_n_trafos.insert(0, str(len(data.get('trafos', []))))
            self.entry_n_ind.delete(0, tk.END); self.entry_n_ind.insert(0, str(len(data.get('inductores', []))))
            self.entry_n_cap.delete(0, tk.END); self.entry_n_cap.insert(0, str(len(data.get('capacitores', []))))
            self.entry_n_cargas.delete(0, tk.END); self.entry_n_cargas.insert(0, str(len(data.get('cargas', []))))
            self.entry_n_gen.delete(0, tk.END); self.entry_n_gen.insert(0, str(len(data.get('generadores', []))))
            self.generar_tablas()
            for i, b in enumerate(data.get('barras', [])):
                if i >= len(self.tipo_barras): break
                self.tipo_barras[i].set(b.get('tipo', 'PQ')); self.V_pref_widgets[i].delete(0, tk.END); self.V_pref_widgets[i].insert(0, b.get('V', '1.0')); self.ang_pref_widgets[i].delete(0, tk.END); self.ang_pref_widgets[i].insert(0, b.get('ang', '0.0'))
            for i, fila in enumerate(data.get('lineas', [])):
                if i >= len(self.entries_lineas): break
                for j, val in enumerate(fila):
                    if j >= len(self.entries_lineas[i]): break
                    self.entries_lineas[i][j].delete(0, tk.END); self.entries_lineas[i][j].insert(0, val)
            for i, fila in enumerate(data.get('trafos', [])):
                if i >= len(self.entries_trafos): break
                for j, val in enumerate(fila):
                    if j >= len(self.entries_trafos[i]): break
                    if j == 4: self.entries_trafos[i][j].set(val)
                    else: self.entries_trafos[i][j].delete(0, tk.END); self.entries_trafos[i][j].insert(0, val)
            for i, fila in enumerate(data.get('inductores', [])):
                if i >= len(self.entries_inductores): break
                for j, val in enumerate(fila): self.entries_inductores[i][j].delete(0, tk.END); self.entries_inductores[i][j].insert(0, val)
            for i, fila in enumerate(data.get('capacitores', [])):
                if i >= len(self.entries_capacitores): break
                for j, val in enumerate(fila): self.entries_capacitores[i][j].delete(0, tk.END); self.entries_capacitores[i][j].insert(0, val)
            for i, fila in enumerate(data.get('cargas', [])): 
                if i >= len(self.entries_cargas): break
                for j, val in enumerate(fila): 
                    self.entries_cargas[i][j].delete(0, tk.END)
                    self.entries_cargas[i][j].insert(0, val)
            for i, fila in enumerate(data.get('generadores', [])):
                if i >= len(self.gen_widgets): break
                for j, val in enumerate(fila): self.gen_widgets[i][j].delete(0, tk.END); self.gen_widgets[i][j].insert(0, val)
            f_data = data.get('falla', {})
            if self.combo_barra_falla: self.combo_barra_falla.set(f_data.get('barra', '1'))
            if self.combo_tipo_falla: self.combo_tipo_falla.set(f_data.get('tipo', 'Monofásica (SLG)'))
            self.entry_ZF.delete(0, tk.END); self.entry_ZF.insert(0, f_data.get('ZF', '0.0'))
            self.entry_Zg.delete(0, tk.END); self.entry_Zg.insert(0, f_data.get('Zg', '0.0'))
            self._dibujar_preview(); self._update_sidebar(); self._set_status(f"Cargado: {os.path.basename(path)}", 'ok')
        except Exception as e: import traceback; traceback.print_exc(); messagebox.showerror("Error al abrir", str(e))

    def _exportar_excel(self):
        if not self.resultados: messagebox.showinfo("Info", "Calcula primero la falla."); return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Libro de Excel", "*.xlsx")])
        if not path: return
        try:
            self._generar_excel(path, self.resultados); self._set_status(f"Excel exportado: {os.path.basename(path)}", 'ok')
            if messagebox.askyesno("Exportado", f"Archivo generado:\n{path}\n\n¿Abrirlo?"):
                try: os.startfile(path)
                except Exception: pass
        except Exception as e: import traceback; traceback.print_exc(); messagebox.showerror("Error", f"No se pudo exportar:\n{e}")

    def _generar_excel(self, ruta, r):
        wb = openpyxl.Workbook(); BASE = r['BASE_MVA']; V_BASE = r['V_BASE_KV']; I_base = BASE / (math.sqrt(3) * V_BASE)
        font_titulo = Font(bold=True, size=14, color="FFFFFF"); fill_titulo = PatternFill("solid", fgColor="1E3A5F")
        head_font = Font(bold=True, color="FFFFFF"); head_fill = PatternFill("solid", fgColor="B91C1C"); center = Alignment(horizontal="center", vertical="center")
        def _titulo(ws, fila, cols, texto):
            c = ws.cell(row=fila, column=1, value=texto); c.font = font_titulo; c.fill = fill_titulo; c.alignment = center
            ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=cols); return fila + 1
        def _headers(ws, fila, headers):
            for j, h in enumerate(headers, start=1):
                c = ws.cell(row=fila, column=j, value=h); c.font = head_font; c.fill = head_fill; c.alignment = center
            return fila + 1
        ws = wb.active; ws.title = "Resumen"
        for col, w in zip("ABCD", [30, 22, 22, 22]): ws.column_dimensions[col].width = w
        f = _titulo(ws, 1, 4, "REPORTE DE FALLA — VATIA ASEP"); f = _headers(ws, f, ["Parámetro", "Valor", "Unidad", "Notas"])
        rows = [("Fecha", datetime.now().strftime("%Y-%m-%d %H:%M"), "", ""), ("Tipo", r['tipo_falla'], "", ""), ("Barra", r['barra_falla'] + 1, "", ""), ("Z_F", r['ZF'], "p.u.", ""), ("Z_g", r['Zg'], "p.u.", ""), ("|V_pref|", abs(r['V_pref_c']), "p.u.", ""), ("Z0_th", abs(r['Zth0']), "p.u.", self._fmt_complex(r['Zth0'])), ("Z1_th", abs(r['Zth1']), "p.u.", self._fmt_complex(r['Zth1'])), ("Z2_th", abs(r['Zth2']), "p.u.", self._fmt_complex(r['Zth2'])), ("I0", abs(r['I0']), "p.u.", self._fmt_complex(r['I0'])), ("I1", abs(r['I1']), "p.u.", self._fmt_complex(r['I1'])), ("I2", abs(r['I2']), "p.u.", self._fmt_complex(r['I2'])), ("Ia", abs(r['Ia']), "p.u.", self._fmt_complex(r['Ia'])), ("Ib", abs(r['Ib']), "p.u.", self._fmt_complex(r['Ib'])), ("Ic", abs(r['Ic']), "p.u.", self._fmt_complex(r['Ic'])), ("Ia", abs(r['Ia']) * I_base, "kA", ""), ("Ib", abs(r['Ib']) * I_base, "kA", ""), ("Ic", abs(r['Ic']) * I_base, "kA", ""), ("S_base", BASE, "MVA", ""), ("V_base", V_BASE, "kV", ""), ("I_base", I_base, "kA", "")]
        for fila in rows:
            for j, v in enumerate(fila, start=1): ws.cell(row=f, column=j, value=v)
            f += 1
        ws2 = wb.create_sheet("Corrientes")
        for col, w in zip("ABCDE", [16, 18, 16, 16, 24]): ws2.column_dimensions[col].width = w
        f2 = _titulo(ws2, 1, 5, f"CORRIENTES — BARRA {r['barra_falla']+1}"); f2 = _headers(ws2, f2, ["Cantidad", "|X| p.u.", "|X| kA", "∠ (°)", "Rectangular"])
        for nombre, val in [("I0", r['I0']), ("I1", r['I1']), ("I2", r['I2']), ("Ia", r['Ia']), ("Ib", r['Ib']), ("Ic", r['Ic'])]:
            ws2.cell(row=f2, column=1, value=nombre); ws2.cell(row=f2, column=2, value=abs(val)); ws2.cell(row=f2, column=3, value=abs(val) * I_base); ws2.cell(row=f2, column=4, value=math.degrees(cmath.phase(val))); ws2.cell(row=f2, column=5, value=self._fmt_complex(val)); f2 += 1
        ws3 = wb.create_sheet("Tensiones")
        for col, w in zip("ABCDEFG", [10, 12, 12, 12, 12, 12, 12]): ws3.column_dimensions[col].width = w
        f3 = _titulo(ws3, 1, 7, "TENSIONES POR BARRA"); f3 = _headers(ws3, f3, ["Barra", "|V0|", "|V1|", "|V2|", "|Va|", "|Vb|", "|Vc|"])
        a_op = complex(-0.5, math.sqrt(3)/2); a2 = a_op**2
        for i in range(r['n']):
            V0 = r['V0_bus'][i]; V1 = r['V1_bus'][i]; V2 = r['V2_bus'][i]; Va = V0 + V1 + V2; Vb = V0 + a2*V1 + a_op*V2; Vc = V0 + a_op*V1 + a2*V2
            for j, v in enumerate([i+1, abs(V0), abs(V1), abs(V2), abs(Va), abs(Vb), abs(Vc)], 1): ws3.cell(row=f3, column=j, value=v)
            f3 += 1
        ws4 = wb.create_sheet("Aportes generadores")
        for col, w in zip("ABCDEFGH", [10]*8): ws4.column_dimensions[col].width = w
        f4 = _titulo(ws4, 1, 8, "APORTES POR GENERADOR"); f4 = _headers(ws4, f4, ["Bus", "|I0|", "|I1|", "|I2|", "|Ia| p.u.", "|Ib| p.u.", "|Ic| p.u.", "|Ia| kA"])
        for a in r.get('aportes', []):
            for j, v in enumerate([a['bus'], abs(a['I0']), abs(a['I1']), abs(a['I2']), abs(a['Ia']), abs(a['Ib']), abs(a['Ic']), abs(a['Ia'])*I_base], 1): ws4.cell(row=f4, column=j, value=v)
            f4 += 1
        ws5 = wb.create_sheet("Ybus-Zbus")
        ws5.column_dimensions['A'].width = 12
        for col in "BCDEFGHIJ": ws5.column_dimensions[col].width = 20
        f5 = 1
        for nombre, M in [("Ybus sec 0", r['Ybus0']), ("Ybus sec 1", r['Ybus1']), ("Ybus sec 2", r['Ybus2']), ("Zbus sec 0", r['Zbus0']), ("Zbus sec 1", r['Zbus1']), ("Zbus sec 2", r['Zbus2'])]:
            f5 = _titulo(ws5, f5, 12, nombre); n = len(M)
            for j in range(n):
                c = ws5.cell(row=f5, column=j+2, value=f"j={j+1}"); c.font = head_font; c.fill = head_fill
            f5 += 1
            for i in range(n):
                ws5.cell(row=f5, column=1, value=f"i={i+1}").font = head_font
                for j in range(n): ws5.cell(row=f5, column=j+2, value=self._fmt_complex(M[i][j], 6))
                f5 += 1
            f5 += 1
        wb.save(ruta)

    def _volver_al_menu(self):
        if not messagebox.askyesno("Confirmar", "¿Volver al menú principal?\nSe perderán los datos no guardados."): return
        try:
            if self._sidebar_after_id: self.root.after_cancel(self._sidebar_after_id); self._sidebar_after_id = None
        except Exception: pass
        for w in self.root.winfo_children(): w.destroy()
        try:
            def abrir_simulador(root):
                root.title("Vatia — Flujos de Potencia (Newton-Raphson)"); root.geometry("1500x900"); root.minsize(1180, 720); root.configure(bg="#EDF1F6")
                app = AppFlujoPotencia(root); root.app = app
            DashboardScreen(self.root, abrir_simulador)
        except Exception: import traceback; traceback.print_exc()

    def _on_close(self):
        try:
            if self._sidebar_after_id: self.root.after_cancel(self._sidebar_after_id); self._sidebar_after_id = None
        except Exception: pass
        try: self.root.destroy()
        except Exception: 
            pass

    def _tab_falla_linea(self, parent):
        r = self.resultados; C = self.C
        n = r['n']
        
        # --- Controles superiores ---
        top = tk.Frame(parent, bg=C['surface'], highlightthickness=1, highlightbackground=C['border'])
        top.pack(fill=tk.X, padx=20, pady=(14, 8))
        
        tk.Label(top, text="Seleccione la Línea:", font=font(10, "bold"), bg=C['surface'], fg=C['primary']).pack(side=tk.LEFT, padx=10, pady=10)
        
        opciones_lineas = [f"L{k+1}: Nodo {L['a']+1} - Nodo {L['b']+1}" for k, L in enumerate(r['lineas'])]
        combo_linea = ttk.Combobox(top, values=opciones_lineas, state="readonly", width=25, font=font(10))
        combo_linea.pack(side=tk.LEFT, padx=10)
        if opciones_lineas: combo_linea.current(0)

        tk.Label(top, text="Tipo de Falla:", font=font(10, "bold"), bg=C['surface'], fg=C['primary']).pack(side=tk.LEFT, padx=(20, 10))
        combo_tipo = ttk.Combobox(top, values=["Monofásica (SLG)", "Trifásica (LLL)"], state="readonly", width=18, font=font(10))
        combo_tipo.pack(side=tk.LEFT, padx=10)
        combo_tipo.current(0)

        tk.Label(top, text="% de la Línea:", font=font(10, "bold"), bg=C['surface'], fg=C['primary']).pack(side=tk.LEFT, padx=(20, 10))
        entry_pct = ttk.Entry(top, width=6, justify='center', font=font(10))
        entry_pct.pack(side=tk.LEFT, padx=10)
        entry_pct.insert(0, "50")

        # --- Área de resultados (scroll) ---
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        canvas.bind('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), 'units'))

        # --- Función que hace el cálculo ---
        def calcular():
            for widget in inner.winfo_children():
                widget.destroy()
                
            idx_linea = combo_linea.current()
            if idx_linea < 0: return
            
            tipo_str = combo_tipo.get()
            pct = self._parse_float(entry_pct.get(), 50.0) / 100.0
            
            # 1. Tomar la línea y partirla en dos
            L = r['lineas'][idx_linea]
            a, b = L['a'], L['b']
            Z1 = L['Z1']; Z2 = L['Z2']; Z0 = L['Z0']; B = L.get('B1_2', 0.0)
            
            # Nodo ficticio es el último (n)
            nuevo_nodo = n
            
            # Creamos las ramas de la línea partida
            L1 = {'a': a, 'b': nuevo_nodo, 'Z1': Z1*pct, 'Z2': Z2*pct, 'Z0': Z0*pct, 'B1_2': B*pct}
            L2 = {'a': nuevo_nodo, 'b': b, 'Z1': Z1*(1-pct), 'Z2': Z2*(1-pct), 'Z0': Z0*(1-pct), 'B1_2': B*(1-pct)}
            
            # Listas temporales con n+1 nodos
            lineas_temp = r['lineas'].copy()
            lineas_temp.pop(idx_linea)
            lineas_temp.extend([L1, L2])
            
            # 2. Construir Ybus temporales para n+1 nodos
            Ybus0 = self._construir_ybus(n+1, lineas_temp, r['trafos'], r.get('inductores',[]), r.get('capacitores',[]), r['generadores'], r.get('cargas',[]), seq=0)
            Ybus1 = self._construir_ybus(n+1, lineas_temp, r['trafos'], r.get('inductores',[]), r.get('capacitores',[]), r['generadores'], r.get('cargas',[]), seq=1)
            Ybus2 = self._construir_ybus(n+1, lineas_temp, r['trafos'], r.get('inductores',[]), r.get('capacitores',[]), r['generadores'], r.get('cargas',[]), seq=2)
            
            Zbus0 = np.linalg.inv(Ybus0); Zbus1 = np.linalg.inv(Ybus1); Zbus2 = np.linalg.inv(Ybus2)
            
            # 3. Impedancias de Thevenin en el nuevo nodo
            Zth0 = Zbus0[nuevo_nodo][nuevo_nodo]
            Zth1 = Zbus1[nuevo_nodo][nuevo_nodo]
            Zth2 = Zbus2[nuevo_nodo][nuevo_nodo]
            
            # Voltaje prefalla (asumimos 1.0 en el punto medio)
            V_pref_c = 1.0 + 0j
            
            # 4. Resolver falla
            if "Monofásica" in tipo_str:
                res = self._resolver_slg(V_pref_c, Zth0, Zth1, Zth2, 0j, 0j)
                tipo_falla = 'SLG'
            else:
                res = self._resolver_lll(V_pref_c, Zth1, 0j)
                tipo_falla = 'LLL'
                
            I0 = res['I0']; I1 = res['I1']; I2 = res['I2']
            
            # 5. Voltajes de falla en todos los nodos (incluyendo el nuevo)
            V0_bus = [-Zbus0[i][nuevo_nodo] * I0 for i in range(n+1)]
            V1_bus = []
            for i in range(n+1):
                Vp_i = 1.0 if i != nuevo_nodo else V_pref_c
                V1_bus.append(Vp_i - Zbus1[i][nuevo_nodo] * I1)
            V2_bus = [-Zbus2[i][nuevo_nodo] * I2 for i in range(n+1)]
            
            # 6. Pasar a ABC
            a_op = complex(-0.5, math.sqrt(3)/2); a2 = a_op ** 2
            Ia = I0 + I1 + I2; Ib = I0 + a2 * I1 + a_op * I2; Ic = I0 + a_op * I1 + a2 * I2
            
            Vabc = []
            for i in range(n+1):
                Va = V0_bus[i] + V1_bus[i] + V2_bus[i]
                Vb = V0_bus[i] + a2*V1_bus[i] + a_op*V2_bus[i]
                Vc = V0_bus[i] + a_op*V1_bus[i] + a2*V2_bus[i]
                Vabc.append((Va, Vb, Vc))
            
            # === IMPRIMIR RESULTADOS ===
            
            # a. Qué hacer
            f = tk.Frame(inner, bg=C['surface'], highlightthickness=1, highlightbackground=C['border'])
            f.pack(fill=tk.X, padx=20, pady=(10, 4))
            tk.Label(f, text="a. PROCEDIMIENTO PARA SIMULAR LA FALLA", font=font(10, "bold"), bg=C['surface'], fg=C['red'], anchor='w').pack(fill=tk.X, padx=14, pady=(8, 2))
            texto_a = f"Se debe crear un nodo ficticio en el punto de falla ({int(pct*100)}% de la línea {a+1}-{b+1}).\nLuego, la línea original se divide en dos ramas: una desde el Nodo {a+1} al nodo ficticio ({n+1}) con impedancia Z*{pct:.2f}, y otra desde el nodo ficticio al Nodo {b+1} con impedancia Z*{1-pct:.2f}.\nFinalmente, se aplica la falla en el nodo ficticio y se recalculan las matrices Ybus y Zbus."
            tk.Label(f, text=texto_a, font=font_m(9), bg=C['surface'], fg=C['text'], justify='left', anchor='w').pack(fill=tk.X, padx=14, pady=(0, 10))

            # b. Corrientes
            BASE = r['BASE_MVA']; V_BASE = r['V_BASE_KV']; I_base = BASE / (math.sqrt(3) * V_BASE)
            self._tabla_resultados(inner, f"b. CORRIENTES DE FALLA EN NODO FICTICIO {n+1} (kA Polar)", ["Componente", "Magnitud (kA)", "Ángulo (°)"],
                [["I0", f"{abs(I0)*I_base:.3f}", f"{math.degrees(cmath.phase(I0)):+.2f}"],
                 ["I1", f"{abs(I1)*I_base:.3f}", f"{math.degrees(cmath.phase(I1)):+.2f}"],
                 ["I2", f"{abs(I2)*I_base:.3f}", f"{math.degrees(cmath.phase(I2)):+.2f}"],
                 ["Ia", f"{abs(Ia)*I_base:.3f}", f"{math.degrees(cmath.phase(Ia)):+.2f}"],
                 ["Ib", f"{abs(Ib)*I_base:.3f}", f"{math.degrees(cmath.phase(Ib)):+.2f}"],
                 ["Ic", f"{abs(Ic)*I_base:.3f}", f"{math.degrees(cmath.phase(Ic)):+.2f}"]], color_col=1)

            # c. Tensiones en los extremos (a y b)
            filas_v = []
            for i in [a, b]:
                V0 = V0_bus[i]; V1 = V1_bus[i]; V2 = V2_bus[i]
                Va, Vb, Vc = Vabc[i]
                filas_v.append([f"Nodo {i+1}", f"{abs(V0):.4f} ∠{math.degrees(cmath.phase(V0)):+.1f}°", f"{abs(V1):.4f} ∠{math.degrees(cmath.phase(V1)):+.1f}°", f"{abs(V2):.4f} ∠{math.degrees(cmath.phase(V2)):+.1f}°", f"{abs(Va):.4f} ∠{math.degrees(cmath.phase(Va)):+.1f}°", f"{abs(Vb):.4f} ∠{math.degrees(cmath.phase(Vb)):+.1f}°", f"{abs(Vc):.4f} ∠{math.degrees(cmath.phase(Vc)):+.1f}°"])
            self._tabla_resultados(inner, "c. TENSIONES EN LOS NODOS EXTREMOS (p.u. Polar)", ["Nodo", "V0", "V1", "V2", "Va", "Vb", "Vc"], filas_v, ancho=18)

            # d. Comparación (Texto didáctico)
            f2 = tk.Frame(inner, bg=C['surface'], highlightthickness=1, highlightbackground=C['border'])
            f2.pack(fill=tk.X, padx=20, pady=(10, 4))
            tk.Label(f2, text="d y e. COMPARACIÓN Y EXPLICACIÓN", font=font(10, "bold"), bg=C['surface'], fg=C['primary'], anchor='w').pack(fill=tk.X, padx=14, pady=(8, 2))
            texto_d = ("Al comparar con la falla en barra, notarás que:\n"
                       "1. La corriente de falla (I) en el punto medio es MENOR. Esto se debe a que la falla está más lejos de los generadores (hay más impedancia de por medio).\n"
                       "2. El voltaje en los nodos extremos (A y B) NO cae tanto a cero como lo haría si la falla fuera directamente en la barra. El voltaje se deprime, pero no colapsa del todo.\n\n"
                       "¿Por qué cambia?\n"
                       "La impedancia equivalente de Thevenin (Zth) en el punto medio de la línea es MAYOR. Al agregar la mitad de la impedancia de la línea, se limita el flujo de corriente de cortocircuito, reduciendo su magnitud y atenuando la caída de tensión en el resto de la red.")
            tk.Label(f2, text=texto_d, font=font_m(9), bg=C['surface'], fg=C['text'], justify='left', anchor='w', wraplength=1000).pack(fill=tk.X, padx=14, pady=(0, 10))

        # Botón para calcular
        btn = tk.Button(top, text="  Calcular Falla en Línea  ", command=calcular, bg=C['red'], fg='white', font=font(10, "bold"), relief=tk.FLAT, padx=16, pady=5, cursor='hand2')
        btn.pack(side=tk.LEFT, padx=20)
        
        # Primera pintura por defecto
        parent.after(100, calcular)

# =====================================================================
# MOTOR DE DESPACHO ECONÓMICO
# =====================================================================
class EconomicDispatchCore:
    """
    Motor de despacho económico con análisis de congestión
    (basado en componentes simétricas DC y factores ISF).
    """

    def __init__(self, n, loads, lines, generators, slack_idx=0):
        """
        n         : número de nodos
        loads     : list[float] — carga en cada nodo [MW]
        lines     : list[dict] con keys {'a','b','x','cap'}
                    a, b → índices 0-based; x en p.u.; cap en MW
        generators: list[dict] con keys {'bus','cap','mc'}
                    bus → índice 0-based; cap en MW; mc en $/MWh
        slack_idx : índice del nodo slack (0-based)
        """
        self.n = n
        self.loads = list(loads)
        self.lines = list(lines)
        self.generators = list(generators)
        self.slack_idx = slack_idx
        self.n_gen = len(self.generators)
        self.n_lin = len(self.lines)

        self._build_B()
        self._build_X()
        self._build_ISF()

    # -----------------------------------------------------------------
    def _build_B(self):
        n = self.n
        B_full = [[0.0]*n for _ in range(n)]
        for L in self.lines:
            a, b, x = L['a'], L['b'], L['x']
            if abs(x) < 1e-12:
                continue
            y = 1.0 / x
            B_full[a][a] += y
            B_full[b][b] += y
            B_full[a][b] -= y
            B_full[b][a] -= y
        self.B_full = B_full

        no_slack = [i for i in range(n) if i != self.slack_idx]
        self.no_slack = no_slack
        m = len(no_slack)
        self.B_red = np.array(
            [[B_full[i][j] for j in no_slack] for i in no_slack],
            dtype=np.float64)

    def _build_X(self):
        n = self.n
        try:
            X_red = np.linalg.inv(self.B_red)
        except np.linalg.LinAlgError:
            raise ValueError("La matriz B reducida es singular "
                             "(red no conectada o slack incorrecto).")
        X_ext = np.zeros((n, n))
        for i, oi in enumerate(self.no_slack):
            for j, oj in enumerate(self.no_slack):
                X_ext[oi][oj] = X_red[i][j]
        self.X_ext = X_ext

    def _build_ISF(self):
        ISF = np.zeros((self.n_lin, self.n))
        for k, L in enumerate(self.lines):
            a, b, x = L['a'], L['b'], L['x']
            if abs(x) < 1e-12:
                continue
            for i in range(self.n):
                ISF[k][i] = (self.X_ext[a][i] - self.X_ext[b][i]) / x
        self.ISF = ISF

    # -----------------------------------------------------------------
    def _compute_flows(self, P_gen):
        """Dado el vector de despacho P_gen (n_gen), retorna (flows, theta)."""
        P_bus = np.zeros(self.n)
        for g in range(self.n_gen):
            P_bus[self.generators[g]['bus']] += P_gen[g]
        for i in range(self.n):
            P_bus[i] -= self.loads[i]
        P_red = np.array([P_bus[i] for i in self.no_slack])
        try:
            theta_red = np.linalg.solve(self.B_red, P_red)
        except np.linalg.LinAlgError:
            return None, None
        theta = np.zeros(self.n)
        for i, oi in enumerate(self.no_slack):
            theta[oi] = theta_red[i]
        flows = np.zeros(self.n_lin)
        for k, L in enumerate(self.lines):
            a, b, x = L['a'], L['b'], L['x']
            if abs(x) < 1e-12:
                continue
            flows[k] = (theta[a] - theta[b]) / x
        return flows, theta

    # -----------------------------------------------------------------
    def _despacho_sin_restricciones(self):
        """Merit order: despacha por costo marginal ascendente."""
        P_total = sum(self.loads)
        orden = sorted(range(self.n_gen),
                       key=lambda g: self.generators[g]['mc'])
        P = np.zeros(self.n_gen)
        rem = P_total
        for g in orden:
            cap = self.generators[g]['cap']
            d = min(cap, rem)
            P[g] = d
            rem -= d
            if rem <= 1e-9:
                break
        return P

    # -----------------------------------------------------------------
    def _redespachar(self, P_inicial, max_iter=30):
        """
        Redespacho iterativo. En cada iteración resuelve la línea más
        sobrecargada ajustando el par (u, d) más económico.
        Retorna (P_final, iteraciones, convergio).
        """
        P = P_inicial.copy()
        convergio = False
        for it in range(max_iter):
            flows, _ = self._compute_flows(P)
            if flows is None:
                return P, it, False

            worst_k, worst_excess = None, 0.0
            for k, L in enumerate(self.lines):
                ex = abs(flows[k]) - L['cap']
                if ex > worst_excess:
                    worst_excess = ex
                    worst_k = k
            if worst_k is None or worst_excess < 1e-3:
                convergio = True
                break

            sign_f = 1.0 if flows[worst_k] > 0 else -1.0
            best = None
            for u in range(self.n_gen):
                bus_u = self.generators[u]['bus']
                avail_up = self.generators[u]['cap'] - P[u]
                if avail_up < 1e-6:
                    continue
                for d in range(self.n_gen):
                    if u == d:
                        continue
                    bus_d = self.generators[d]['bus']
                    avail_down = P[d]
                    if avail_down < 1e-6:
                        continue
                    isf_diff = self.ISF[worst_k][bus_u] - self.ISF[worst_k][bus_d]
                    delta_flow = sign_f * isf_diff
                    if delta_flow >= -1e-9:
                        continue
                    delta_needed = worst_excess / (-delta_flow)
                    delta = min(delta_needed, avail_up, avail_down)
                    if delta < 1e-9:
                        continue
                    cost_delta = (self.generators[u]['mc']
                                  - self.generators[d]['mc']) * delta
                    if best is None or cost_delta < best[2]:
                        best = (u, d, cost_delta, delta)
            if best is None:
                break
            u, d, _, delta = best
            P[u] += delta
            P[d] -= delta
        return P, it + 1, convergio

    # -----------------------------------------------------------------
    def _precio_nodal(self, k, P_final, flows_final):
        """
        Precio marginal nodal en el bus k: mínimo costo de servir 1 MW
        adicional respetando las restricciones activas.
        Plan: subir gen u por 'a' MW, bajar gen d por (a-1) MW.
        Balance: a - (a-1) = 1  ✓
        """
        congest = [ℓ for ℓ in range(self.n_lin)
                   if abs(flows_final[ℓ]) >= self.lines[ℓ]['cap'] - 1e-3]

        # Sin congestión → precio uniforme
        if not congest:
            for g in range(self.n_gen):
                if (P_final[g] > 1e-6 and
                    P_final[g] < self.generators[g]['cap'] - 1e-6):
                    return self.generators[g]['mc']
            return max(g['mc'] for g in self.generators)

        best_cost = None

        # Caso A: solo subir un generador en 1 MW
        for u in range(self.n_gen):
            bus_u = self.generators[u]['bus']
            avail_up = self.generators[u]['cap'] - P_final[u]
            if avail_up < 1.0 - 1e-6:
                continue
            ok = True
            for ℓ in congest:
                sign_f = 1.0 if flows_final[ℓ] > 0 else -1.0
                dF = self.ISF[ℓ][bus_u] - self.ISF[ℓ][k]
                if sign_f * dF > 1e-9:
                    ok = False
                    break
            if ok:
                cost = self.generators[u]['mc']
                if best_cost is None or cost < best_cost:
                    best_cost = cost

        # Caso B: subir u por a, bajar d por (a-1)
        for u in range(self.n_gen):
            bus_u = self.generators[u]['bus']
            avail_up = self.generators[u]['cap'] - P_final[u]
            if avail_up < 1.0 - 1e-6:
                continue
            for d in range(self.n_gen):
                if u == d:
                    continue
                bus_d = self.generators[d]['bus']
                avail_down = P_final[d]
                if avail_down < 1e-6:
                    continue

                a_lo, a_hi = 1.0, min(avail_up, avail_down + 1.0)
                if a_lo > a_hi + 1e-9:
                    continue

                feasible = True
                for ℓ in congest:
                    sign_f = 1.0 if flows_final[ℓ] > 0 else -1.0
                    A = sign_f * (self.ISF[ℓ][bus_u] - self.ISF[ℓ][bus_d])
                    B = sign_f * (self.ISF[ℓ][bus_d] - self.ISF[ℓ][k])
                    # A*a + B ≤ 0
                    if abs(A) < 1e-9:
                        if B > 1e-9:
                            feasible = False
                            break
                    elif A > 0:
                        a_hi = min(a_hi, -B / A)
                    else:
                        a_lo = max(a_lo, -B / A)
                if not feasible or a_lo > a_hi + 1e-9:
                    continue

                if self.generators[u]['mc'] > self.generators[d]['mc']:
                    a_opt = a_lo
                else:
                    a_opt = a_hi

                cost = (a_opt * self.generators[u]['mc']
                        - (a_opt - 1) * self.generators[d]['mc'])
                if cost < 0:
                    cost = max(self.generators[u]['mc'],
                               self.generators[d]['mc'])
                if best_cost is None or cost < best_cost:
                    best_cost = cost

        if best_cost is None:
            return max(g['mc'] for g in self.generators)
        return max(best_cost, 0.0)

    def descomponer_flujos(self, P_gen):
        """
        Descompone los flujos por superposición:
          F_ℓ = Σ_g ISF_ℓ,bus_g · P_g  −  Σ_i ISF_ℓ,i · L_i
        Retorna:
          contrib_gen  : matriz (n_lin × n_gen)
          contrib_load : matriz (n_lin × n)
          contrib_neta : matriz (n_lin × n)  (gen − load por bus)
        """
        n_l = self.n_lin
        n_g = self.n_gen
        n = self.n

        contrib_gen = np.zeros((n_l, n_g))
        for ℓ in range(n_l):
            for g in range(n_g):
                bus_g = self.generators[g]['bus']
                contrib_gen[ℓ][g] = self.ISF[ℓ][bus_g] * P_gen[g]

        contrib_load = np.zeros((n_l, n))
        for ℓ in range(n_l):
            for i in range(n):
                contrib_load[ℓ][i] = -self.ISF[ℓ][i] * self.loads[i]

        contrib_neta = np.zeros((n_l, n))
        for ℓ in range(n_l):
            for i in range(n):
                inyeccion_neta = -self.loads[i]
                for g in range(n_g):
                    if self.generators[g]['bus'] == i:
                        inyeccion_neta += P_gen[g]
                contrib_neta[ℓ][i] = self.ISF[ℓ][i] * inyeccion_neta

        return contrib_gen, contrib_load, contrib_neta
            # Descomposición por superposición
        contrib_gen, contrib_load, contrib_neta = \
            self.descomponer_flujos(P_fin)
    def calcular_alternativas_redespacho(self, P_ED, flows_ED):
        """
        Encuentra TODAS las alternativas viables de redespacho que resuelven
        las congestiones. Ordenadas por costo total ascendente.
        """
        alternativas = []
        congestionadas = [k for k, L in enumerate(self.lines)
                          if abs(flows_ED[k]) > L['cap'] + 1e-3]
        if not congestionadas:
            return alternativas

        # Generar candidatos (u, d) por cada línea congestionada
        candidatos = set()
        for k in congestionadas:
            sign_f = 1.0 if flows_ED[k] > 0 else -1.0
            excess = abs(flows_ED[k]) - self.lines[k]['cap']
            for u in range(self.n_gen):
                bus_u = self.generators[u]['bus']
                avail_up = self.generators[u]['cap'] - P_ED[u]
                if avail_up < 1e-6:
                    continue
                for d in range(self.n_gen):
                    if u == d:
                        continue
                    bus_d = self.generators[d]['bus']
                    avail_down = P_ED[d]
                    if avail_down < 1e-6:
                        continue
                    isf_diff = self.ISF[k][bus_u] - self.ISF[k][bus_d]
                    delta_flow = sign_f * isf_diff
                    if delta_flow >= -1e-9:
                        continue
                    delta_needed = excess / (-delta_flow)
                    delta0 = min(delta_needed, avail_up, avail_down)
                    if delta0 < 1e-6:
                        continue
                    candidatos.add((u, d, k))

        # Para cada candidato, aplicar y luego iterar hasta resolver todo
        for (u, d, k0) in candidatos:
            sign_f = 1.0 if flows_ED[k0] > 0 else -1.0
            excess = abs(flows_ED[k0]) - self.lines[k0]['cap']
            bus_u = self.generators[u]['bus']
            bus_d = self.generators[d]['bus']
            isf_diff = self.ISF[k0][bus_u] - self.ISF[k0][bus_d]
            if sign_f * isf_diff >= -1e-9:
                continue
            delta0 = min(excess / (-sign_f * isf_diff),
                         self.generators[u]['cap'] - P_ED[u],
                         P_ED[d])
            if delta0 < 1e-6:
                continue
            P_init = P_ED.copy()
            P_init[u] += delta0
            P_init[d] -= delta0
            # Redespachar a partir de aquí
            P_fin, iters, conv = self._redespachar(P_init)
            flows_fin, _ = self._compute_flows(P_fin)
            if flows_fin is None:
                continue
            costo = sum(self.generators[g]['mc'] * P_fin[g]
                        for g in range(self.n_gen))
            resuelve = all(abs(flows_fin[kk]) <= self.lines[kk]['cap'] + 1e-3
                           for kk in range(self.n_lin))

            # Evitar duplicados con mismo despacho
            dup = False
            for a in alternativas:
                if all(abs(a['P'][g] - P_fin[g]) < 0.01
                       for g in range(self.n_gen)):
                    dup = True
                    break
            if dup:
                continue

            alternativas.append({
                'P': P_fin.tolist(),
                'flows': flows_fin.tolist(),
                'costo': costo,
                'delta_inicial': delta0,
                'gen_up': u,
                'gen_down': d,
                'linea_base': k0,
                'iteraciones': iters,
                'resuelve_todo': resuelve,
            })

        alternativas.sort(key=lambda a: a['costo'])
        return alternativas

    # -----------------------------------------------------------------
    def calcular(self):
        """Ejecuta el pipeline completo y devuelve alternativas."""
        # 1) Despacho sin restricciones
        P_ED = self._despacho_sin_restricciones()
        flows_ED, _ = self._compute_flows(P_ED)
        costo_ED = sum(self.generators[g]['mc'] * P_ED[g]
                       for g in range(self.n_gen))

        # 2) Detectar congestión
        congestionadas = [k for k, L in enumerate(self.lines)
                          if abs(flows_ED[k]) > L['cap'] + 1e-6]

        # 3) Redespacho óptimo
        if congestionadas:
            P_fin, iters, conv = self._redespachar(P_ED)
            alternativas = self.calcular_alternativas_redespacho(P_ED, flows_ED)
        else:
            P_fin, iters, conv = P_ED.copy(), 0, True
            alternativas = []

        flows_fin, _ = self._compute_flows(P_fin)
        costo_fin = sum(self.generators[g]['mc'] * P_fin[g]
                        for g in range(self.n_gen))

        # 4) Precios nodales
        precios = [self._precio_nodal(k, P_fin, flows_fin)
                   for k in range(self.n)]
        
        # Descomposición por superposición
        contrib_gen, contrib_load, contrib_neta = \
            self.descomponer_flujos(P_fin)

        return {
            'n': self.n,
            'loads': self.loads,
            'lines': self.lines,
            'generators': self.generators,
            'P_ED': P_ED.tolist(),
            'flows_ED': flows_ED.tolist(),
            'costo_ED': costo_ED,
            'P_final': P_fin.tolist(),
            'flows_final': flows_fin.tolist(),
            'costo_final': costo_fin,
            'costo_seguridad': costo_fin - costo_ED,
            'congestionadas': congestionadas,
            'iteraciones': iters,
            'convergio': conv,
            'precios_nodales': precios,
            'ISF': self.ISF.tolist(),
            'alternativas': alternativas,   # ← NUEVO
            'contrib_gen': contrib_gen.tolist(),
            'contrib_load': contrib_load.tolist(),
            'contrib_neta': contrib_neta.tolist(),
        }

# =====================================================================
# MÓDULO — DESPACHO ECONÓMICO (UI completa)
# =====================================================================
class EconomicDispatchScreen:
    """Módulo de Despacho Económico con análisis de congestión."""

    C = {
        'bg':          '#EDF1F6',
        'panel':       '#FFFFFF',
        'surface':     '#F7F9FC',
        'sidebar':     '#E2E8F0',
        'primary':     '#1E3A5F',
        'primary_l':   '#2C5282',
        'primary_d':   '#152741',
        'accent':      '#0E7490',
        'accent_l':    '#0891B2',
        'slack':       '#047857',
        'pv':          '#6D28D9',
        'pq':          '#C2410C',
        'info':        '#0369A1',
        'amber':       '#B45309',
        'purple':      '#6D28D9',
        'text':        '#1F2937',
        'text_l':      '#4B5563',
        'text_ll':     '#94A3B8',
        'border':      '#CBD5E1',
        'head_bg':     '#E2E8F0',
        'row_even':    '#FFFFFF',
        'row_odd':     '#F7F9FC',
        'success':     '#047857',
        'warning':     '#B45309',
        'danger':      '#B91C1C',
        'gray_dark':   '#475569',
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Vatia — Despacho Económico")
        self.root.geometry("1500x900")
        self.root.minsize(1180, 720)
        self.root.configure(bg=self.C['bg'])

        self.resultados = None
        self.n_gen = 3 
        self.load_widgets = []
        self.line_widgets = []
        self.gen_widgets = []
        self.sidebar_labels = {}
        self._sidebar_after_id = None

        self._build_menu()
        self._build_header()
        self._build_toolbar()
        self._build_main_area()
        self._build_action_bar()
        self._build_status_bar()

        self.root.bind('<Control-n>', lambda e: self._nuevo_proyecto())
        self.root.bind('<Control-s>', lambda e: self._guardar_proyecto())
        self.root.bind('<Control-o>', lambda e: self._abrir_proyecto())
        self.root.bind('<F5>', lambda e: self.calcular_despacho())
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.generar_tablas()

    # =================================================================
    # HELPERS BÁSICOS
    # =================================================================
    @staticmethod
    def _parse_float(s, default=0.0):
        try:
            if s is None: return default
            s = str(s).strip().replace(',', '.')
            return float(s) if s else default
        except Exception:
            return default

    @staticmethod
    def _parse_int(s, default=0):
        try:
            if s is None: return default
            s = str(s).strip().replace(',', '.')
            return int(float(s)) if s else default
        except Exception:
            return default

    def _set_status(self, text, kind='info'):
        try:
            self.status_var.set(text)
            colors = {'info': self.C['accent'], 'ok': self.C['success'],
                      'warn': self.C['warning'], 'err': self.C['danger']}
            self.status_icon.config(fg=colors.get(kind, self.C['accent']))
        except Exception:
            pass

    # =================================================================
    # MENÚ
    # =================================================================
    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        m_arch = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Archivo", menu=m_arch)
        m_arch.add_command(label="Nuevo proyecto\t\tCtrl+N",
                            command=self._nuevo_proyecto)
        m_arch.add_command(label="Abrir proyecto...\tCtrl+O",
                            command=self._abrir_proyecto)
        m_arch.add_command(label="Guardar\t\tCtrl+S",
                            command=self._guardar_proyecto)
        m_arch.add_separator()
        m_arch.add_command(label="Exportar a Excel...",
                            command=self._exportar_excel)
        m_arch.add_separator()
        m_arch.add_command(label="Volver al menú",
                            command=self._volver_al_menu)

        m_calc = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Cálculo", menu=m_calc)
        m_calc.add_command(label="Calcular despacho\tF5",
                            command=self.calcular_despacho)
        m_calc.add_separator()
        m_calc.add_command(label="Ver resultados",
                            command=self.mostrar_resultados)

        m_ayuda = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Ayuda", menu=m_ayuda)
        m_ayuda.add_command(label="Teoría del método",
                            command=self.mostrar_teoria)

    # =================================================================
    # HEADER
    # =================================================================
    def _build_header(self):
        C = self.C
        hdr = tk.Frame(self.root, bg=C['primary_d'], height=90)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['purple'], height=4).pack(fill=tk.X, side=tk.TOP)

        left = tk.Frame(hdr, bg=C['primary_d'])
        left.pack(side=tk.LEFT, padx=24, pady=14)

        cv = tk.Canvas(left, width=54, height=54, bg=C['primary_d'],
                       highlightthickness=0)
        cv.pack(side=tk.LEFT, padx=(0, 14))
        cv.create_oval(2, 2, 52, 52, fill=C['purple'], outline='')
        cv.create_text(27, 27, text="$", font=font(18, "bold"), fill='white')

        tinfo = tk.Frame(left, bg=C['primary_d'])
        tinfo.pack(side=tk.LEFT)
        tk.Label(tinfo, text="DESPACHO ECONÓMICO",
                 font=font(15, "bold"), bg=C['primary_d'], fg='white'
                 ).pack(anchor='w')
        tk.Label(tinfo,
                 text="Merit order  ·  Congestión  ·  Precios nodales",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(anchor='w')

        right = tk.Frame(hdr, bg=C['primary_d'])
        right.pack(side=tk.RIGHT, padx=24, pady=15)
        badge = tk.Frame(right, bg=C['purple'])
        badge.pack()
        tk.Label(badge, text="  MÉTODO DC  ",
                 font=font(10, "bold"), bg=C['purple'], fg='white',
                 padx=10, pady=7).pack()

    # =================================================================
    # TOOLBAR
    # =================================================================
    def _build_toolbar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(10, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['border'])
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel'])
        inner.pack(padx=18, pady=12, fill=tk.X)

        # --- Número de nodos ---
        tk.Label(inner, text="Número de Nodos (N):", font=font(10, "bold"),
                 bg=C['panel'], fg=C['text']).pack(side=tk.LEFT)
        self.entry_n = ttk.Entry(inner, width=6, justify='center',
                                  font=font(11))
        self.entry_n.pack(side=tk.LEFT, padx=8, ipady=3)
        self.entry_n.insert(0, "3")

        # --- Número de generadores ---
        tk.Label(inner, text="Número de Generadores:",
                 font=font(10, "bold"),
                 bg=C['panel'], fg=C['text']).pack(side=tk.LEFT, padx=(20, 0))
        self.entry_n_gen = ttk.Entry(inner, width=6, justify='center',
                                       font=font(11))
        self.entry_n_gen.pack(side=tk.LEFT, padx=8, ipady=3)
        self.entry_n_gen.insert(0, "3")

        # --- Botón generar ---
        tk.Button(inner, text="  Generar Tablas  ",
                  command=self.generar_tablas,
                  bg=C['primary'], fg='white',
                  font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                  padx=16, pady=7, borderwidth=0
                  ).pack(side=tk.LEFT, padx=10)

        tk.Button(inner, text="  Nuevo  ", command=self._nuevo_proyecto,
                  bg=C['surface'], fg=C['text'],
                  font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                  padx=14, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=2)

        tk.Button(inner, text="  Abrir  ", command=self._abrir_proyecto,
                  bg=C['surface'], fg=C['text'],
                  font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                  padx=14, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=2)

        tk.Button(inner, text="  Guardar  ", command=self._guardar_proyecto,
                  bg=C['surface'], fg=C['text'],
                  font=font(10, "bold"), relief=tk.FLAT, cursor='hand2',
                  padx=14, pady=7, borderwidth=0).pack(side=tk.LEFT, padx=2)

        tk.Button(inner, text="◄  Volver al menú",
                  command=self._volver_al_menu,
                  bg=C['surface'], fg=C['primary'],
                  font=font(10, "bold"), relief=tk.FLAT,
                  cursor='hand2', padx=16, pady=7, borderwidth=0
                  ).pack(side=tk.RIGHT, padx=2)

    # =================================================================
    # MAIN AREA
    # =================================================================
    def _build_main_area(self):
        C = self.C
        main = tk.Frame(self.root, bg=C['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)

        self.paned = ttk.PanedWindow(main, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        self.sidebar = tk.Frame(self.paned, bg=C['sidebar'], width=210)
        self.paned.add(self.sidebar, weight=0)
        self._build_sidebar()

        center = tk.Frame(self.paned, bg=C['bg'])
        self.paned.add(center, weight=5)
        self.notebook = ttk.Notebook(center)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        right = tk.Frame(self.paned, bg=C['bg'], width=380)
        self.paned.add(right, weight=4)
        right_lf = tk.Frame(right, bg=C['panel'], highlightthickness=1,
                             highlightbackground=C['border'])
        right_lf.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        tk.Label(right_lf, text="  VISTA PREVIA DEL SISTEMA  ",
                 font=font(10, "bold"), bg=C['head_bg'], fg=C['purple'],
                 pady=8).pack(fill=tk.X)

        self.preview_canvas = tk.Canvas(right_lf, bg=C['panel'],
                                          highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.preview_canvas.bind('<Configure>',
                                  lambda e: self._dibujar_preview())

    # =================================================================
    # SIDEBAR
    # =================================================================
    def _build_sidebar(self):
        C = self.C
        sb = self.sidebar
        sb.pack_propagate(False)

        tk.Label(sb, text="  PANEL DE CONTROL  ",
                 font=font(10, "bold"), bg=C['purple'], fg='white',
                 pady=10).pack(fill=tk.X)
        tk.Frame(sb, bg=C['accent'], height=2).pack(fill=tk.X)

        def _stat(key, label, color=None):
            frame = tk.Frame(sb, bg=C['sidebar'])
            frame.pack(fill=tk.X, padx=14, pady=(8, 2))
            tk.Label(frame, text=label.upper(), font=font(8, "bold"),
                     bg=C['sidebar'], fg=C['text_l']).pack(anchor='w')
            lbl = tk.Label(frame, text="—", font=font(13, "bold"),
                           bg=C['sidebar'], fg=color or C['text'],
                           anchor='w', justify='left', wraplength=180)
            lbl.pack(anchor='w', fill=tk.X)
            self.sidebar_labels[key] = lbl

        def _sep():
            tk.Frame(sb, bg=C['border'], height=1).pack(
                fill=tk.X, padx=14, pady=6)

        _stat('nodos', 'Nodos', C['accent'])
        _stat('lineas', 'Líneas', C['info'])
        _stat('generadores', 'Generadores', C['pv'])
        _sep()
        _stat('carga', 'Carga total (MW)', C['pq'])
        _stat('costo_ed', 'Costo sin restricciones', C['slack'])
        _stat('costo_fin', 'Costo con congestión', C['danger'])
        _stat('cs', 'Costo de seguridad', C['warning'])
        _sep()
        _stat('estado', 'Estado', C['text_l'])

        pie = tk.Frame(sb, bg=C['sidebar'])
        pie.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        tk.Label(pie, text="por Kevin", font=font(9, italic=True),
                 bg=C['sidebar'], fg=C['text_l']).pack()
        tk.Label(pie, text="Vatia © 2026", font=font_m(8),
                 bg=C['sidebar'], fg=C['text_ll']).pack()

        self.root.after(400, self._update_sidebar)

    def _update_sidebar(self):
        try:
            if not self.root.winfo_exists():
                return
        except Exception:
            return

        if not getattr(self, 'sidebar_labels', None):
            try:
                self._sidebar_after_id = self.root.after(
                    500, self._update_sidebar)
            except Exception:
                pass
            return

        try:
            if not self.entry_n.winfo_exists():
                return
        except Exception:
            return

        C = self.C                                      # ← NUEVA LÍNEA

        try:
            n = int(self.entry_n.get())
            n_lineas = sum(1 for row in self.line_widgets
                           if self._parse_int(row[0].get(), 0) > 0
                           and self._parse_int(row[1].get(), 0) > 0)
            n_gen = sum(1 for g in self.gen_widgets if abs(self._parse_float(g[2].get(), 0.0)) > 1e-9)
            carga = sum(self._parse_float(w.get(), 0.0)
                        for w in self.load_widgets)

            self.sidebar_labels['nodos'].config(text=str(n))
            self.sidebar_labels['lineas'].config(text=str(n_lineas))
            self.sidebar_labels['generadores'].config(text=str(n_gen))
            self.sidebar_labels['carga'].config(text=f"{carga:.1f}")

            if self.resultados:
                r = self.resultados
                self.sidebar_labels['costo_ed'].config(
                    text=f"${r['costo_ED']:.2f}/h")
                self.sidebar_labels['costo_fin'].config(
                    text=f"${r['costo_final']:.2f}/h")
                self.sidebar_labels['cs'].config(
                    text=f"${r['costo_seguridad']:.2f}/h")
                if r['congestionadas']:
                    self.sidebar_labels['estado'].config(
                        text="CON CONGESTIÓN", fg=C['warning'])
                else:
                    self.sidebar_labels['estado'].config(
                        text="ÓPTIMO", fg=C['success'])
            else:
                for k in ('costo_ed', 'costo_fin', 'cs'):
                    self.sidebar_labels[k].config(text="—")
                self.sidebar_labels['estado'].config(
                    text="SIN CALCULAR", fg=C['text_l'])
        except tk.TclError:
            return
        except Exception:
            pass

        try:
            if self.root.winfo_exists():
                self._sidebar_after_id = self.root.after(
                    500, self._update_sidebar)
        except Exception:
            pass

    # =================================================================
    # ACTION BAR
    # =================================================================
    def _build_action_bar(self):
        C = self.C
        outer = tk.Frame(self.root, bg=C['bg'])
        outer.pack(fill=tk.X, padx=14, pady=(6, 4))
        card = tk.Frame(outer, bg=C['panel'], highlightthickness=1,
                        highlightbackground=C['border'])
        card.pack(fill=tk.X)
        inner = tk.Frame(card, bg=C['panel'])
        inner.pack(padx=14, pady=10, fill=tk.X)

        def _grupo(parent, titulo):
            g = tk.Frame(parent, bg=C['panel'])
            g.pack(side=tk.LEFT, padx=(0, 20))
            tk.Label(g, text=titulo.upper(), font=font(8, "bold"),
                     bg=C['panel'], fg=C['text_l']).pack(anchor='w', pady=(0, 4))
            fila = tk.Frame(g, bg=C['panel'])
            fila.pack()
            return fila

        g1 = _grupo(inner, "Cálculo")
        btn = tk.Button(g1, text="  ▶  Calcular Despacho  ",
                  command=self.calcular_despacho,
                  bg=C['purple'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=20, pady=7,
                  borderwidth=0,
                  activebackground=C['primary_l'], activeforeground='white')
        btn.pack(side=tk.LEFT, padx=2)
        ToolTip(btn, "Ejecuta el despacho económico con análisis de congestión.")

        g2 = _grupo(inner, "Vista")
        btn_r = tk.Button(g2, text="  Ver Resultados  ",
                  command=self.mostrar_resultados,
                  bg=C['accent'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=16, pady=7,
                  borderwidth=0)
        btn_r.pack(side=tk.LEFT, padx=2)
        ToolTip(btn_r, "Abre la ventana de resultados.")

        btn_t = tk.Button(g2, text="  Teoría  ",
                  command=self.mostrar_teoria,
                  bg=C['info'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=16, pady=7,
                  borderwidth=0)
        btn_t.pack(side=tk.LEFT, padx=2)
        ToolTip(btn_t, "Muestra la teoría del despacho económico.")

        g3 = tk.Frame(inner, bg=C['panel'])
        g3.pack(side=tk.RIGHT)
        tk.Label(g3, text="ARCHIVO", font=font(8, "bold"),
                 bg=C['panel'], fg=C['text_l']).pack(anchor='e', pady=(0, 4))
        fila3 = tk.Frame(g3, bg=C['panel'])
        fila3.pack()
        tk.Button(fila3, text="  Exportar Excel  ",
                  command=self._exportar_excel,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, cursor='hand2', padx=16, pady=7,
                  borderwidth=0).pack(side=tk.LEFT, padx=2)

    # =================================================================
    # STATUS BAR
    # =================================================================
    def _build_status_bar(self):
        C = self.C
        sb = tk.Frame(self.root, bg=C['primary_d'], height=32)
        sb.pack(fill=tk.X, side=tk.BOTTOM)
        sb.pack_propagate(False)
        tk.Frame(sb, bg=C['purple'], width=3).pack(side=tk.LEFT, fill=tk.Y)
        self.status_icon = tk.Label(sb, text="●", font=font(11),
                                     bg=C['primary_d'], fg=C['success'])
        self.status_icon.pack(side=tk.LEFT, padx=(12, 4))
        self.status_var = tk.StringVar(value="Listo.")
        tk.Label(sb, textvariable=self.status_var, font=font(9),
                 bg=C['primary_d'], fg='white').pack(side=tk.LEFT)
        tk.Label(sb, text="v1.0  ·  Despacho",
                 font=font_m(8), bg=C['primary_d'],
                 fg=C['text_ll']).pack(side=tk.RIGHT, padx=(0, 14))

    def _crear_tab_generadores(self, n_gen):
        """Crea la pestaña de generadores con n_gen filas independientes
        del número de nodos."""
        C = self.C
        hdr = tk.Frame(self.tab_gens, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"GENERADORES ({n_gen})",
                 font=font(12, "bold"), bg=C['primary_d'], fg='white',
                 pady=12).pack()
        tk.Label(hdr,
                 text="Nodo de conexión · Capacidad (MW) · Costo marginal ($/MWh)",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(self.tab_gens, bg=C['panel'])
        grid.pack(padx=20, pady=20, fill=tk.X)

        headers = ["Gen #", "Nodo", "Capacidad (MW)",
                   "Costo marginal ($/MWh)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"),
                     bg=C['head_bg'], fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)

        for i in range(n_gen):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']

            # Etiqueta Gen #
            tk.Label(grid, text=f"G{i+1}", font=font(10, "bold"),
                     bg=bg, fg=C['purple'], padx=8, pady=6
                     ).grid(row=i+1, column=0, padx=1, pady=1, sticky='nsew')

            # 3 entries: nodo, capacidad, costo marginal
            row = []
            for j in range(3):
                e = ttk.Entry(grid, justify='center', font=font(9))
                e.grid(row=i+1, column=j+1, padx=1, pady=1, ipady=4,
                       sticky='nsew')
                e.bind('<KeyRelease>', lambda ev: self._update_sidebar())
                row.append(e)
            self.gen_widgets.append(row)

        # Nota informativa
        note = tk.Frame(self.tab_gens, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        note.pack(fill=tk.X, padx=20, pady=(0, 20))
        tk.Label(note,
                 text="Nota:  Puede haber varios generadores en el mismo nodo. "
                      "El nodo de conexión debe estar entre 1 y N. "
                      "Los generadores con capacidad 0 se ignoran en el cálculo.",
                 font=font_m(9), bg=C['surface'], fg=C['text_l'],
                 justify='left', anchor='w', wraplength=1000
                 ).pack(fill=tk.X, padx=14, pady=10)

    # =================================================================
    # GENERAR TABLAS
    # =================================================================
    def generar_tablas(self):
        # --- Leer N ---
        try:
            n = int(self.entry_n.get())
            if n < 2:
                messagebox.showwarning("Advertencia", "Mínimo 2 nodos.")
                return
            if n > 20:
                messagebox.showwarning("Advertencia", "Máximo 20 nodos.")
                return
        except ValueError:
            messagebox.showwarning("Advertencia",
                "Ingrese un número entero en Nodos.")
            return

        # --- Leer N_gen ---
        try:
            n_gen = int(self.entry_n_gen.get())
            if n_gen < 1:
                messagebox.showwarning("Advertencia",
                    "Debe haber al menos 1 generador.")
                return
            if n_gen > 40:
                messagebox.showwarning("Advertencia",
                    "Máximo 40 generadores.")
                return
        except ValueError:
            messagebox.showwarning("Advertencia",
                "Ingrese un número entero en Generadores.")
            return

        self.n_gen = n_gen

        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)

        self.load_widgets.clear()
        self.line_widgets.clear()
        self.gen_widgets.clear()
        self.resultados = None

        self.tab_cargas = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_lineas = tk.Frame(self.notebook, bg=self.C['panel'])
        self.tab_gens = tk.Frame(self.notebook, bg=self.C['panel'])
        self.notebook.add(self.tab_cargas, text='   Cargas (nodos)   ')
        self.notebook.add(self.tab_lineas, text='   Líneas (x, cap)   ')
        self.notebook.add(self.tab_gens,
                          text=f'   Generadores ({n_gen})   ')

        self._crear_tab_cargas(n)
        self._crear_tab_lineas(n)
        self._crear_tab_generadores(n_gen)    # ← usamos n_gen, no n

        self._set_status(
            f"Sistema: {n} nodos, {n_gen} generadores configurados.", 'info')
        self._update_sidebar()
        self.root.after(120, self._dibujar_preview)

    # =================================================================
    # TABS DE ENTRADA
    # =================================================================
    def _crear_tab_cargas(self, n):
        C = self.C
        hdr = tk.Frame(self.tab_cargas, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="CARGAS POR NODO",
                 font=font(12, "bold"), bg=C['primary_d'], fg='white',
                 pady=12).pack()
        tk.Label(hdr, text="Demanda en MW en cada nodo",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(self.tab_cargas, bg=C['panel'])
        grid.pack(padx=20, pady=20, fill=tk.X)

        headers = ["Nodo", "Carga (MW)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"),
                     bg=C['head_bg'], fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)

        for i in range(n):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']
            tk.Label(grid, text=str(i+1), font=font(10, "bold"),
                     bg=bg, fg=C['purple'], padx=8, pady=6
                     ).grid(row=i+1, column=0, padx=1, pady=1, sticky='nsew')
            e = ttk.Entry(grid, justify='center', font=font(9))
            e.grid(row=i+1, column=1, padx=1, pady=1, ipady=4, sticky='nsew')
            e.insert(0, "0.0")
            e.bind('<KeyRelease>', lambda ev: self._update_sidebar())
            self.load_widgets.append(e)

    def _crear_tab_lineas(self, n):
        C = self.C
        hdr = tk.Frame(self.tab_lineas, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="LÍNEAS DE TRANSMISIÓN",
                 font=font(12, "bold"), bg=C['primary_d'], fg='white',
                 pady=12).pack()
        tk.Label(hdr, text="Reactancia (p.u.) y capacidad (MW)",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(self.tab_lineas, bg=C['panel'])
        grid.pack(padx=20, pady=20, fill=tk.X)

        headers = ["Nodo A", "Nodo B", "x (p.u.)", "Capacidad (MW)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"),
                     bg=C['head_bg'], fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)

        for i in range(n):
            row = []
            for j in range(4):
                e = ttk.Entry(grid, justify='center', font=font(9))
                e.grid(row=i+1, column=j, padx=1, pady=1, ipady=4,
                       sticky='nsew')
                e.bind('<KeyRelease>', lambda ev: self._update_sidebar())
                row.append(e)
            self.line_widgets.append(row)

    def _crear_tab_generadores(self, n_gen):
        C = self.C
        hdr = tk.Frame(self.tab_gens, bg=C['primary_d'])
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text=f"GENERADORES ({n_gen})",
                 font=font(12, "bold"), bg=C['primary_d'], fg='white',
                 pady=12).pack()
        tk.Label(hdr,
                 text="Nodo de conexión · Capacidad (MW) · Costo marginal ($/MWh)",
                 font=font(9, italic=True), bg=C['primary_d'],
                 fg=C['text_l']).pack(pady=(0, 10))

        grid = tk.Frame(self.tab_gens, bg=C['panel'])
        grid.pack(padx=20, pady=20, fill=tk.X)

        headers = ["Gen #", "Nodo", "Capacidad (MW)",
                   "Costo marginal ($/MWh)"]
        for i, h in enumerate(headers):
            tk.Label(grid, text=h, font=font(9, "bold"),
                     bg=C['head_bg'], fg=C['accent'], padx=8, pady=7
                     ).grid(row=0, column=i, padx=1, pady=1, sticky='nsew')
            grid.columnconfigure(i, weight=1)

        for i in range(n_gen):
            bg = C['row_even'] if i % 2 == 0 else C['row_odd']
            tk.Label(grid, text=f"G{i+1}", font=font(10, "bold"),
                     bg=bg, fg=C['purple'], padx=8, pady=6
                     ).grid(row=i+1, column=0, padx=1, pady=1, sticky='nsew')
            row = []
            for j in range(3):
                e = ttk.Entry(grid, justify='center', font=font(9))
                e.grid(row=i+1, column=j+1, padx=1, pady=1, ipady=4,
                       sticky='nsew')
                e.bind('<KeyRelease>', lambda ev: self._update_sidebar())
                row.append(e)
            self.gen_widgets.append(row)

        note = tk.Frame(self.tab_gens, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        note.pack(fill=tk.X, padx=20, pady=(0, 20))
        tk.Label(note,
                 text="Nota:  Puede haber varios generadores en el mismo nodo. "
                      "El nodo de conexión debe estar entre 1 y N. "
                      "Los generadores con capacidad 0 se ignoran.",
                 font=font_m(9), bg=C['surface'], fg=C['text_l'],
                 justify='left', anchor='w', wraplength=1000
                 ).pack(fill=tk.X, padx=14, pady=10)

    def _dibujar_preview(self):
        canvas = getattr(self, 'preview_canvas', None)
        if canvas is None:
            return
        try:
            n = int(self.entry_n.get())
        except Exception:
            return
        if n < 2:
            return
        canvas.delete("all")
        W = canvas.winfo_width(); H = canvas.winfo_height()
        if W < 50 or H < 50:
            W, H = 520, 520
        C = self.C
        canvas.create_rectangle(0, 0, W, H, fill='#FFFFFF', outline='')
        canvas.create_text(W/2, 26, text="Vista previa — Despacho",
                           font=font(13, "bold"), fill=C['purple'])
        canvas.create_line(W/2-100, 42, W/2+100, 42,
                           fill=C['purple'], width=2)

        lineas = []
        for row in self.line_widgets:
            a = self._parse_int(row[0].get(), 0)
            b = self._parse_int(row[1].get(), 0)
            if a > 0 and b > 0 and a <= n and b <= n and a != b:
                lineas.append((a-1, b-1))

        if not lineas:
            canvas.create_text(W/2, H/2,
                text="Sin líneas definidas",
                font=font(11), fill=C['text_ll'])
            return

        cx, cy = W/2, H/2 + 20
        R = min(W, H) * 0.30
        pos = []
        for i in range(n):
            ang = 2 * math.pi * i / n - math.pi/2
            pos.append((cx + R*math.cos(ang), cy + R*math.sin(ang)))

        for (a, b) in lineas:
            x1, y1 = pos[a]; x2, y2 = pos[b]
            canvas.create_line(x1, y1, x2, y2, width=2.5,
                               fill=C['gray_dark'])

        for i, (x, y) in enumerate(pos):
            canvas.create_oval(x-15, y-15, x+15, y+15,
                               fill=C['purple'], outline='white', width=2)
            canvas.create_text(x, y, text=str(i+1), fill='white',
                               font=font(10, "bold"))
    # =================================================================
    # PREVIEW (simple)
    # =================================================================
    def _dibujar_preview(self):
        canvas = getattr(self, 'preview_canvas', None)
        if canvas is None:
            return
        try:
            n = int(self.entry_n.get())
        except Exception:
            return
        if n < 2:
            return
        canvas.delete("all")
        W = canvas.winfo_width(); H = canvas.winfo_height()
        if W < 50 or H < 50:
            W, H = 520, 520
        C = self.C
        canvas.create_rectangle(0, 0, W, H, fill='#FFFFFF', outline='')
        canvas.create_text(W/2, 26, text="Vista previa — Despacho",
                           font=font(13, "bold"), fill=C['purple'])
        canvas.create_line(W/2-100, 42, W/2+100, 42,
                           fill=C['purple'], width=2)

        lineas = []
        for row in self.line_widgets:
            a = self._parse_int(row[0].get(), 0)
            b = self._parse_int(row[1].get(), 0)
            if a > 0 and b > 0 and a <= n and b <= n and a != b:
                lineas.append((a-1, b-1))

        if not lineas:
            canvas.create_text(W/2, H/2,
                text="Sin líneas definidas",
                font=font(11), fill=C['text_ll'])
            return

        cx, cy = W/2, H/2 + 20
        R = min(W, H) * 0.30
        pos = []
        for i in range(n):
            ang = 2 * math.pi * i / n - math.pi/2
            pos.append((cx + R*math.cos(ang), cy + R*math.sin(ang)))

        for (a, b) in lineas:
            x1, y1 = pos[a]; x2, y2 = pos[b]
            canvas.create_line(x1, y1, x2, y2, width=2.5,
                               fill=C['gray_dark'])

        for i, (x, y) in enumerate(pos):
            canvas.create_oval(x-15, y-15, x+15, y+15,
                               fill=C['purple'], outline='white', width=2)
            canvas.create_text(x, y, text=str(i+1), fill='white',
                               font=font(10, "bold"))

    # =================================================================
    # CÁLCULO
    # =================================================================
    def calcular_despacho(self):
        try:
            n = int(self.entry_n.get())
            loads = [self._parse_float(w.get(), 0.0)
                     for w in self.load_widgets]

            lines = []
            for row in self.line_widgets:
                a = self._parse_int(row[0].get(), 0) - 1
                b = self._parse_int(row[1].get(), 0) - 1
                x = self._parse_float(row[2].get(), 0.1)
                cap = self._parse_float(row[3].get(), 100.0)
                if 0 <= a < n and 0 <= b < n and a != b and abs(x) > 1e-9:
                    lines.append({'a': a, 'b': b, 'x': x, 'cap': cap})

            gens = []
            for row in self.gen_widgets:
                bus = self._parse_int(row[0].get(), 0) - 1
                cap = self._parse_float(row[1].get(), 0.0)
                mc = self._parse_float(row[2].get(), 0.0)
                if 0 <= bus < n and cap > 0:
                    gens.append({'bus': bus, 'cap': cap, 'mc': mc})

            if not lines or not gens:
                messagebox.showerror("Error",
                    "Defina al menos una línea y un generador.")
                return

            core = EconomicDispatchCore(n, loads, lines, gens, slack_idx=0)
            self._core = core                       # ← guardar referencia
            self.resultados = core.calcular()
            self._set_status("Despacho calculado.", 'ok')
            self._update_sidebar()
            self.mostrar_resultados()
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Error en el cálculo:\n{e}")

    # =================================================================
    # MOSTRAR RESULTADOS (ventana con pestañas)
    # =================================================================
    def mostrar_resultados(self):
        if not self.resultados:
            messagebox.showinfo("Info", "Calcula primero el despacho.")
            return
        r = self.resultados
        C = self.C

        win = tk.Toplevel(self.root)
        win.title("Resultados — Despacho Económico")
        win.geometry("1250x820")
        win.minsize(1050, 680)
        win.configure(bg=C['bg'])

        hdr = tk.Frame(win, bg=C['primary_d'], height=70)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Frame(hdr, bg=C['purple'], height=4).pack(fill=tk.X, side=tk.TOP)
        tk.Label(hdr, text="  RESULTADOS — DESPACHO ECONÓMICO",
                 font=font(13, "bold"), bg=C['primary_d'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=22)
        tk.Button(hdr, text="✕", command=win.destroy,
                  bg=C['danger'], fg='white', font=font(11, "bold"),
                  relief=tk.FLAT, padx=14, pady=6, cursor='hand2'
                  ).pack(side=tk.RIGHT, padx=12, pady=22)

        nb = ttk.Notebook(win)
        nb.pack(fill='both', expand=True, padx=14, pady=10)

        # --- Pestaña 1: Resumen ---
        t1 = tk.Frame(nb, bg=C['bg'])
        nb.add(t1, text='Resumen')
        self._tab_resumen(t1, r)

        # --- Pestaña 2: Despacho ---
        t2 = tk.Frame(nb, bg=C['bg'])
        nb.add(t2, text='Despacho')
        self._tab_despacho(t2, r)

        # --- Pestaña 3: Flujos ---
        t3 = tk.Frame(nb, bg=C['bg'])
        nb.add(t3, text='Flujos')
        self._tab_flujos(t3, r)

        # --- Pestaña 4: Precios nodales ---
        t4 = tk.Frame(nb, bg=C['bg'])
        nb.add(t4, text='Precios nodales')
        self._tab_precios(t4, r)

        # --- Pestaña 5: ISF ---
        t5 = tk.Frame(nb, bg=C['bg'])
        nb.add(t5, text='ISF')
        self._tab_isf(t5, r)

        # --- Pestaña 6: Esquemático ---
        t6 = tk.Frame(nb, bg=C['bg'])
        nb.add(t6, text='Esquemático')
        self._tab_redespachos(t6, r)

        # --- Pestaña 7: Redespachos ---
        t7 = tk.Frame(nb, bg=C['bg'])
        nb.add(t7, text='Redespachos')
        self._tab_redespachos(t7, r)

        # --- Pestaña 8: Paso a paso ---
        t8 = tk.Frame(nb, bg=C['bg'])
        nb.add(t8, text='Paso a paso')
        self._tab_paso_a_paso(t8, r)

        # --- Pestaña 9: Contribuciones ---
        t9 = tk.Frame(nb, bg=C['bg'])
        nb.add(t9, text='Contribuciones')
        self._tab_contribuciones(t9, r)
    
        tk.Button(win, text="Cerrar", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=24, pady=8, cursor='hand2'
                  ).pack(pady=8)
        # =================================================================
    # PESTAÑA REDESPACHOS (con selector + diagramas unifilares)
    # =================================================================
    def _tab_redespachos(self, parent, r):
        C = self.C
        n = r['n']
        gens = r['generators']
        loads = r['loads']
        lines = r['lines']
        P_ED = r['P_ED']
        flows_ED = r['flows_ED']
        costo_ED = r['costo_ED']
        alternativas = r.get('alternativas', [])

        # === Toolbar con selector ===
        top = tk.Frame(parent, bg=C['surface'], highlightthickness=1,
                       highlightbackground=C['border'])
        top.pack(fill=tk.X, padx=20, pady=(14, 8))

        tk.Label(top, text="ALTERNATIVAS DE REDESPACHO",
                 font=font(11, "bold"), bg=C['surface'], fg=C['purple'],
                 padx=14, pady=8).pack(side=tk.LEFT)

        if not alternativas:
            tk.Label(top, text="   (Sistema sin congestión — no hay redespachos)",
                     font=font(10, italic=True), bg=C['surface'],
                     fg=C['success'], pady=8).pack(side=tk.LEFT)

        # === Opciones del combo: ED + cada alternativa ===
        opciones = [f"Sin redespacho (ED puro) → ${costo_ED:.2f}/h"]
        for i, alt in enumerate(alternativas):
            tag = " ✓" if alt['resuelve_todo'] else " ✗"
            opciones.append(
                f"Redespacho {i+1}: G{alt['gen_up']+1}↑ / G{alt['gen_down']+1}↓ "
                f"→ ${alt['costo']:.2f}/h{tag}")
        self._redesp_var = tk.StringVar(value=opciones[0])
        combo = ttk.Combobox(top, textvariable=self._redesp_var,
                             values=opciones, state="readonly",
                             width=58, font=font(9))
        combo.pack(side=tk.LEFT, padx=10, pady=8)

        # === Área principal: diagrama + tabla comparativa ===
        main = tk.Frame(parent, bg=C['bg'])
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 14))

        # --- Diagrama (izquierda, grande) ---
        left = tk.Frame(main, bg=C['bg'])
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        cv_w, cv_h = 1000, 700      # ← antes era 780, 560
        cvs = tk.Canvas(left, bg='#FFFFFF', highlightthickness=1,
                        highlightbackground=C['border'],
                        width=cv_w, height=cv_h,
                        scrollregion=(0, 0, cv_w, cv_h))
        cvs.pack(fill=tk.BOTH, expand=True)

        # --- Tabla comparativa (derecha) ---
        right = tk.Frame(main, bg=C['bg'], width=380)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(14, 0))
        right.pack_propagate(False)

        # === Encabezado de la comparativa ===
        tk.Label(right, text="COMPARATIVA DE OPCIONES",
                 font=font(10, "bold"), bg=C['purple'], fg='white',
                 pady=8).pack(fill=tk.X)
        comp = tk.Frame(right, bg=C['border'])
        comp.pack(fill=tk.X, pady=(0, 8))

        # Header
        head = tk.Frame(comp, bg=C['primary_d'])
        head.pack(fill=tk.X)
        for h, w in [("Opción", 22), ("Costo $/h", 10), ("Δ", 8)]:
            tk.Label(head, text=h, font=font(9, "bold"),
                     bg=C['primary_d'], fg='white', width=w,
                     pady=6).pack(side=tk.LEFT)

        # Filas
        def _fila(texto, costo, color):
            f = tk.Frame(comp, bg=C['border'])
            f.pack(fill=tk.X, pady=1)
            tk.Label(f, text=texto, font=font_m(8),
                     bg=C['row_even'], fg=C['text'], width=22,
                     anchor='w', pady=5).pack(side=tk.LEFT)
            tk.Label(f, text=f"${costo:.2f}", font=font_m(9, "bold"),
                     bg=C['row_even'], fg=color, width=10,
                     pady=5).pack(side=tk.LEFT)
            delta = costo - costo_ED
            tk.Label(f, text=f"+{delta:.2f}" if delta > 0 else "base",
                     font=font_m(8), bg=C['row_even'],
                     fg=C['warning'] if delta > 0 else C['success'],
                     width=8, pady=5).pack(side=tk.LEFT)

        _fila("ED puro (sin restr.)", costo_ED, C['success'])
        for i, alt in enumerate(alternativas):
            color = C['purple'] if alt['resuelve_todo'] else C['danger']
            _fila(f"Redesp. {i+1}: G{alt['gen_up']+1}↑/G{alt['gen_down']+1}↓",
                  alt['costo'], color)

        # === Leyenda ===
        ley = tk.Frame(right, bg=C['surface'], highlightthickness=1,
                       highlightbackground=C['border'])
        ley.pack(fill=tk.X, pady=(8, 0))
        tk.Label(ley, text="Notas:", font=font(8, "bold"),
                 bg=C['surface'], fg=C['text'], anchor='w'
                 ).pack(fill=tk.X, padx=10, pady=(6, 2))
        tk.Label(ley, text="✓ = resuelve todas las congestiones",
                 font=font(8), bg=C['surface'], fg=C['success'],
                 anchor='w').pack(fill=tk.X, padx=10)
        tk.Label(ley, text="✗ = aún quedan líneas sobrecargadas",
                 font=font(8), bg=C['surface'], fg=C['danger'],
                 anchor='w').pack(fill=tk.X, padx=10)
        tk.Label(ley, text="↑ = generador que sube potencia",
                 font=font(8), bg=C['surface'], fg=C['text_l'],
                 anchor='w').pack(fill=tk.X, padx=10)
        tk.Label(ley, text="↓ = generador que baja potencia",
                 font=font(8), bg=C['surface'], fg=C['text_l'],
                 anchor='w').pack(fill=tk.X, padx=10, pady=(0, 8))

        # === Función para redibujar ===
        def _redibujar(event=None):
            idx = combo.current()
            cvs.delete("all")
            if idx <= 0:
                P = P_ED
                flows = flows_ED
                costo = costo_ED
                titulo = "DESPACHO SIN RESTRICCIONES (ED PURO)"
                self._dibujar_unifilar_despacho(
                    cvs, n, gens, loads, lines, P, flows, costo,
                    titulo, cv_w, cv_h, con_redespacho=False,
                    P_base=None,
                    precios=r['precios_nodales'])
            else:
                alt = alternativas[idx - 1]
                P = alt['P']
                flows = alt['flows']
                costo = alt['costo']
                titulo = (f"REDESPACHO {idx}: "
                          f"G{alt['gen_up']+1} ↑ / G{alt['gen_down']+1} ↓")
                self._dibujar_unifilar_despacho(
                    cvs, n, gens, loads, lines, P, flows, costo,
                    titulo, cv_w, cv_h, con_redespacho=True,
                    P_base=P_ED,               # ← marca los ΔP
                    precios=r['precios_nodales'])
            _resaltar(idx)

        def _resaltar(idx):
            for child in comp.winfo_children()[1:]:
                for sub in child.winfo_children():
                    sub.config(bg=C['row_even'])
            children = comp.winfo_children()[1:]
            if 0 <= idx < len(children):
                for sub in children[idx].winfo_children():
                    sub.config(bg='#FEF3C7')

        combo.bind('<<ComboboxSelected>>', _redibujar)

        # Primera pintura
        parent.after(50, _redibujar)

    # =================================================================
    # DIBUJO DE UNIFILAR PARA DESPACHO
    # =================================================================
    def _dibujar_unifilar_despacho(self, canvas, n, gens, loads, lines,
                                    P, flows, costo, titulo, W, H,
                                    con_redespacho=False, P_base=None,
                                    precios=None):
        C = self.C
        canvas.delete("all")

        # ─── Bandas reservadas (comprimidas verticalmente) ───
        y_title    = 22
        y_subtitle = 44
        y_gen_lbl  = 78      # etiqueta arriba del generador
        y_gen      = 130     # centro del círculo G
        y_arc      = 180     # peak de arcos (más arriba para no chocar)
        y_node     = 330     # centro de barras
        y_pi       = 362     # π nodal (debajo del nodo)
        y_load     = 470     # base del triángulo (punta)
        y_load_lbl = 515     # etiqueta de carga

        # ─── Grid sutil ───
        for x in range(0, W, 50):
            canvas.create_line(x, 0, x, H, fill='#F8FAFC')
        for y in range(0, H, 50):
            canvas.create_line(0, y, W, y, fill='#F8FAFC')

        # ─── Título ───
        color_titulo = C['purple'] if con_redespacho else C['success']
        canvas.create_text(W / 2, y_title, text=titulo,
                           font=font(14, "bold"), fill=color_titulo)
        canvas.create_text(
            W / 2, y_subtitle,
            text=f"Costo = ${costo:.2f}/h    ·    "
                 f"Gen = {sum(P):.1f} MW    ·    "
                 f"Carga = {sum(loads):.1f} MW",
            font=font(9, italic=True), fill=C['text_l'])

        # ─── Columnas de nodos ───
        margin = 140
        usable = W - 2 * margin
        col_x = ([W / 2] if n == 1 else
                 [margin + i * usable / (n - 1) for i in range(n)])

        # ─── LÍNEAS ───
        cong_set = set(k for k, L in enumerate(lines)
                       if abs(flows[k]) > L['cap'] + 1e-3)
        
        labels_lines = []  # Guardar etiquetas para dibujarlas al final

        for k, L in enumerate(lines):
            a, b = L['a'], L['b']
            xa, xb = col_x[a], col_x[b]
            es_cong = k in cong_set
            color = C['danger'] if es_cong else C['info']
            lw = 4 if es_cong else 3
            F = flows[k]
            estado = "  ⚠" if es_cong else ""
            
            # Formato compacto en 2 líneas (mucho más angosto)
            label = (f"L{k+1}   F={F:+.1f} MW{estado}\n"
                     f"x={L['x']:.3f} · cap={L['cap']:.0f} MW")
            bg_lbl = '#FEE2E2' if es_cong else '#FFFFFF'

            if abs(a - b) == 1:
                # Línea adyacente (recta)
                direction = 1 if xb > xa else -1
                sx1, sx2 = xa + direction * 34, xb - direction * 34
                canvas.create_line(sx1, y_node, sx2, y_node,
                                   fill=color, width=lw, capstyle=tk.ROUND)

                fwd = F >= 0
                aux = (1 if fwd else -1) * (1 if xb > xa else -1)
                px, py = 0, 1
                sz, hw = 18, 10
                mid = (sx1 + sx2) / 2
                canvas.create_polygon(
                    mid + aux * sz / 2, y_node,
                    mid - aux * sz / 2 + px * hw, y_node + py * hw,
                    mid - aux * sz / 2 - px * hw, y_node - py * hw,
                    fill=color, outline='white', width=1.5)

                # Etiqueta PEGADA a la línea horizontal
                lx = mid
                ly = y_node - 26
                labels_lines.append((lx, ly, label, color, bg_lbl))
            else:
                # Línea no adyacente (arco)
                y_arc_local = y_arc - 18 * max(0, abs(a - b) - 1)
                xa_dir = xa + (34 if xb > xa else -34)
                xb_dir = xb + (-34 if xb > xa else 34)
                mid_x = (xa + xb) / 2

                pts = []
                n_seg = 40
                for j in range(n_seg + 1):
                    t = j / n_seg
                    bx = (1-t)**2 * xa_dir + 2*(1-t)*t * mid_x + t**2 * xb_dir
                    by = (1-t)**2 * y_node + 2*(1-t)*t * y_arc_local + t**2 * y_node
                    pts.extend([bx, by])
                canvas.create_line(*pts, fill=color, width=lw,
                                   smooth=True, capstyle=tk.ROUND)

                ta = 0.5
                bx = (1-ta)**2 * xa_dir + 2*(1-ta)*ta * mid_x + ta**2 * xb_dir
                by = (1-ta)**2 * y_node + 2*(1-ta)*ta * y_arc_local + ta**2 * y_node
                dx_dt = 2*(1-ta)*(mid_x - xa_dir) + 2*ta*(xb_dir - mid_x)
                dy_dt = 2*(1-ta)*(y_arc_local - y_node) + 2*ta*(y_node - y_arc_local)
                alen = math.hypot(dx_dt, dy_dt) or 1.0
                fwd = F >= 0
                ux = (dx_dt / alen) if fwd else (-dx_dt / alen)
                uy = (dy_dt / alen) if fwd else (-dy_dt / alen)
                px, py = -uy, ux
                sz, hw = 18, 10
                canvas.create_polygon(
                    bx + ux * sz / 2, by + uy * sz / 2,
                    bx - ux * sz / 2 + px * hw, by - uy * sz / 2 + py * hw,
                    bx - ux * sz / 2 - px * hw, by - uy * sz / 2 - py * hw,
                    fill=color, outline='white', width=1.5)

                # Etiqueta movida a la izquierda y pegada al arco
                lx, ly = mid_x - 80, y_arc_local - 5
                labels_lines.append((lx, ly, label, color, bg_lbl))

        # ─── DIBUJAR ETIQUETAS DE LÍNEAS ───
        for lx, ly, label, color, bg_lbl in labels_lines:
            tid = canvas.create_text(lx, ly, text=label,
                                     font=font_m(8, "bold"),
                                     fill=color, justify=tk.CENTER)
            bb = canvas.bbox(tid)
            if bb:
                # Un padding más pequeño y borde fino
                canvas.create_rectangle(bb[0] - 5, bb[1] - 2,
                                         bb[2] + 5, bb[3] + 2,
                                         fill=bg_lbl, outline=color, width=1)
                canvas.tag_raise(tid)

        # ─── GENERADORES ───
        gens_por_bus = {}
        for gi, g in enumerate(gens):
            gens_por_bus.setdefault(g['bus'], []).append((gi, g))

        r_gen = 22
        for bus, lista in gens_por_bus.items():
            x_node = col_x[bus]
            k = len(lista)
            conv_y = y_node - 50

            if k == 1:
                gen_pos = [(x_node, y_gen)]
            else:
                spacing = min(170, 780 / k)
                total_w = (k - 1) * spacing
                start_x = x_node - total_w / 2
                gen_pos = [(start_x + j * spacing, y_gen) for j in range(k)]

            for (gi, g), (gx, gy) in zip(lista, gen_pos):
                activo = P[gi] > 1e-3
                col = C['purple'] if activo else '#94A3B8'

                canvas.create_line(gx, gy + r_gen + 4,
                                   x_node, conv_y,
                                   width=2,
                                   fill=C['gray_dark'] if activo else '#CBD5E1',
                                   dash=None if activo else (4, 3))
                canvas.create_line(x_node, conv_y,
                                   x_node, y_node - 30,
                                   width=2.5, fill=C['gray_dark'])

                canvas.create_oval(gx - r_gen, gy - r_gen,
                                   gx + r_gen, gy + r_gen,
                                   fill=col, outline='white', width=3)
                canvas.create_text(gx, gy - 3, text="G",
                                   font=font(12, "bold"), fill='white')
                canvas.create_text(gx, gy + 8,
                                   text=f"${g['mc']:.1f}",
                                   font=font(8, "bold"), fill='white')

                lbl = (f"G{gi+1}   P={P[gi]:.1f} MW\n"
                       f"cap={g['cap']:.0f}   MC={g['mc']:.2f}")
                tid = canvas.create_text(gx, y_gen_lbl, text=lbl,
                                         font=font_m(8, "bold"),
                                         fill=col, justify=tk.CENTER)
                bb = canvas.bbox(tid)
                if bb:
                    canvas.create_rectangle(bb[0] - 5, bb[1] - 2,
                                             bb[2] + 5, bb[3] + 2,
                                             fill='white', outline=col,
                                             width=1)
                    canvas.tag_raise(tid)

                if P_base is not None:
                    delta = P[gi] - P_base[gi]
                    if abs(delta) > 1e-3:
                        col_d = '#047857' if delta > 0 else '#B91C1C'
                        signo = "↑" if delta > 0 else "↓"
                        txt_d = f"{signo} {delta:+.1f} MW"
                        
                        # Poner al lado del círculo
                        if gx > W/2:
                            lx_d, anchor_d = gx - r_gen - 10, 'e'
                        else:
                            lx_d, anchor_d = gx + r_gen + 10, 'w'
                            
                        tid2 = canvas.create_text(
                            lx_d, gy, text=txt_d, anchor=anchor_d,
                            font=font(9, "bold"), fill=col_d)
                        bb2 = canvas.bbox(tid2)
                        if bb2:
                            canvas.create_rectangle(
                                bb2[0] - 4, bb2[1] - 2,
                                bb2[2] + 4, bb2[3] + 2,
                                fill='white', outline=col_d, width=1)
                            canvas.tag_raise(tid2)

        # ─── NODOS ───
        for i in range(n):
            x = col_x[i]
            canvas.create_line(x, y_node - 32, x, y_node + 32,
                               width=14, fill='#1F2937', capstyle=tk.ROUND)
            canvas.create_line(x, y_node - 32, x, y_node + 32,
                               width=10, fill=C['amber'], capstyle=tk.ROUND)
            canvas.create_oval(x - 22, y_node - 22, x + 22, y_node + 22,
                               fill=C['amber'], outline='white', width=3)
            canvas.create_text(x, y_node, text=str(i + 1),
                               font=font(14, "bold"), fill='white')

            if precios is not None:
                tid = canvas.create_text(
                    x, y_pi,
                    text=f"π{i+1} = ${precios[i]:.2f}/MWh",
                    font=font(9, "bold"), fill=C['purple'])
                bb = canvas.bbox(tid)
                if bb:
                    canvas.create_rectangle(
                        bb[0] - 5, bb[1] - 2, bb[2] + 5, bb[3] + 2,
                        fill='#F3E8FF', outline=C['purple'], width=1)
                    canvas.tag_raise(tid)

        # ─── CARGAS ───
        for i in range(n):
            if loads[i] <= 0:
                continue
            x = col_x[i]

            canvas.create_line(x, y_node + 32, x, y_load - 26,
                               width=2.5, fill=C['gray_dark'])

            canvas.create_polygon(
                x - 26, y_load - 26,
                x + 26, y_load - 26,
                x, y_load,
                fill=C['pq'], outline='#7C2D12', width=2)

            tid = canvas.create_text(x, y_load_lbl,
                                     text=f"D{i+1} = {loads[i]:.1f} MW",
                                     font=font(10, "bold"), fill=C['pq'])
            bb = canvas.bbox(tid)
            if bb:
                canvas.create_rectangle(bb[0] - 8, bb[1] - 4,
                                         bb[2] + 8, bb[3] + 4,
                                         fill='#FFF7ED', outline=C['pq'],
                                         width=1)
                canvas.tag_raise(tid)

        # ─── Nota al pie ───
        canvas.create_text(20, H - 12,
                           text="↑ sube   ·   ↓ baja   ·   ⚠ línea congestionada",
                           font=font(8, italic=True), fill=C['text_l'],
                           anchor='w')
    # =================================================================
    # CONSTRUCCIÓN DE LA SECUENCIA DE PASOS
    # =================================================================
    def _construir_pasos(self, r):
        """Arma la lista de pasos pedagógicos para la pestaña 'Paso a paso'."""
        n_gen = len(r['generators'])
        n_lin = len(r['lines'])
        pasos = []

        # ---- Paso 1: Sistema base ----
        pasos.append({
            'titulo': 'PASO 1 — SISTEMA ORIGINAL',
            'subtitulo': 'Cargas, generadores y líneas definidas. Aún sin despacho.',
            'descripcion': (
                f"• {r['n']} nodos, {n_gen} generadores, {n_lin} líneas\n"
                f"• Carga total: {sum(r['loads']):.1f} MW\n"
                f"• No hay flujos ni generación asignada todavía.\n"
                f"• Cada generador tiene su costo marginal ($/MWh) y "
                f"cada línea su reactancia (p.u.) y capacidad (MW)."),
            'P': [0.0] * n_gen,
            'flows': [0.0] * n_lin,
            'costo': 0.0,
            'P_base': None,
            'precios': None,
            'con_redespacho': False,
        })

        # ---- Paso 2: ED puro (merit order) ----
        pasos.append({
            'titulo': 'PASO 2 — DESPACHO ECONÓMICO SIN RESTRICCIONES',
            'subtitulo': 'Merit order: se despacha del generador más barato al más caro.',
            'descripcion': (
                f"• Se ordenan los generadores por costo marginal ascendente.\n"
                f"• Se despachan hasta cubrir la demanda total "
                f"({sum(r['loads']):.1f} MW).\n"
                f"• Costo total sin restricciones: ${r['costo_ED']:.2f}/h\n"
                f"• Este es el óptimo económico pero NO considera la red."),
            'P': r['P_ED'],
            'flows': r['flows_ED'],
            'costo': r['costo_ED'],
            'P_base': None,
            'precios': None,
            'con_redespacho': False,
        })

        # ---- Paso 3: Verificación de flujos ----
        cong = r['congestionadas']
        if cong:
            lineas_cong = ", ".join(f"L{k+1}" for k in cong)
            desc_ver = (
                f"• Se calculan los flujos DC por cada línea.\n"
                f"• ⚠ Líneas congestionadas detectadas: {lineas_cong}\n"
                f"• El ED puro NO es viable porque viola la capacidad "
                f"de esas líneas.\n"
                f"• Se necesita REDESPACHAR.")
        else:
            desc_ver = (
                f"• Todos los flujos cumplen |F| ≤ capacidad.\n"
                f"• ✓ El ED puro es factible. No se requiere redespacho.\n"
                f"• El costo de seguridad es $0.")
        pasos.append({
            'titulo': 'PASO 3 — VERIFICACIÓN DE FLUJOS Y CONGESTIÓN',
            'subtitulo': 'Se comprueba si algún flujo excede la capacidad.',
            'descripcion': desc_ver,
            'P': r['P_ED'],
            'flows': r['flows_ED'],
            'costo': r['costo_ED'],
            'P_base': None,
            'precios': None,
            'con_redespacho': False,
        })

        # ---- Pasos 4..N: cada redespacho ----
        for i, alt in enumerate(r['alternativas']):
            tag = "✓ resuelve todas las congestiones" \
                  if alt['resuelve_todo'] else "✗ quedan sobrecargas"
            cs = alt['costo'] - r['costo_ED']
            desc = (
                f"• Estrategia: subir G{alt['gen_up']+1} y bajar "
                f"G{alt['gen_down']+1}.\n"
                f"• ΔP inicial: {alt['delta_inicial']:.2f} MW "
                f"(para aliviar L{alt['linea_base']+1}).\n"
                f"• Iteraciones internas: {alt['iteraciones']}.\n"
                f"• {tag}\n"
                f"• Costo restringido: ${alt['costo']:.2f}/h\n"
                f"• Costo de seguridad: +${cs:.2f}/h sobre el ED puro.")
            pasos.append({
                'titulo': (f'PASO {4+i} — ALTERNATIVA {i+1}: '
                           f'G{alt["gen_up"]+1}↑ / G{alt["gen_down"]+1}↓'),
                'subtitulo': ('Redespacho óptimo considerando las restricciones '
                              'de las líneas.'),
                'descripcion': desc,
                'P': alt['P'],
                'flows': alt['flows'],
                'costo': alt['costo'],
                'P_base': r['P_ED'],
                'precios': None,
                'con_redespacho': True,
            })

        # ---- Paso final: precios nodales ----
        precios = r['precios_nodales']
        desc_pr = (
            "• Precio marginal nodal = costo de servir 1 MW extra en cada "
            "nodo, respetando las restricciones activas.\n"
            + "".join(f"   · π{k+1} = ${precios[k]:.2f}/MWh\n"
                      for k in range(r['n']))
            + f"• Costo total del despacho final: ${r['costo_final']:.2f}/h\n"
            f"• Costo de seguridad: ${r['costo_seguridad']:.2f}/h")
        pasos.append({
            'titulo': f'PASO {4+len(r["alternativas"])} — PRECIOS NODALES',
            'subtitulo': 'Precio marginal de la energía en cada nodo.',
            'descripcion': desc_pr,
            'P': r['P_final'],
            'flows': r['flows_final'],
            'costo': r['costo_final'],
            'P_base': None,
            'precios': precios,
            'con_redespacho': False,
        })

        return pasos

    # =================================================================
    # PESTAÑA PASO A PASO
    # =================================================================
    def _tab_paso_a_paso(self, parent, r):
        core_global = self._core
        C = self.C
        pasos = self._construir_pasos(r)
        estado = {'idx': 0}

        # ─── Barra superior: título + navegación ───
        top = tk.Frame(parent, bg=C['surface'], highlightthickness=1,
                       highlightbackground=C['border'])
        top.pack(fill=tk.X, padx=20, pady=(14, 8))

        lbl_titulo = tk.Label(top, text="", font=font(12, "bold"),
                              bg=C['surface'], fg=C['purple'],
                              anchor='w')
        lbl_titulo.pack(side=tk.LEFT, padx=14, pady=10)

        # Botones de navegación (a la derecha)
        nav = tk.Frame(top, bg=C['surface'])
        nav.pack(side=tk.RIGHT, padx=10, pady=6)

        btn_prev = tk.Button(nav, text="  ◄  Anterior  ",
                             bg=C['primary'], fg='white',
                             font=font(9, "bold"), relief=tk.FLAT,
                             cursor='hand2', padx=12, pady=6,
                             borderwidth=0)
        btn_prev.pack(side=tk.LEFT, padx=2)
        btn_next = tk.Button(nav, text="  Siguiente  ► ",
                             bg=C['primary'], fg='white',
                             font=font(9, "bold"), relief=tk.FLAT,
                             cursor='hand2', padx=12, pady=6,
                             borderwidth=0)
        btn_next.pack(side=tk.LEFT, padx=2)
        lbl_ind = tk.Label(nav, text="", font=font(9, "bold"),
                           bg=C['surface'], fg=C['text'], width=10)
        lbl_ind.pack(side=tk.LEFT, padx=8)

        # ─── Canvas del diagrama ───
        cv_w, cv_h = 1000, 620
        cvs_wrap = tk.Frame(parent, bg=C['bg'])
        cvs_wrap.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 6))
        cvs = tk.Canvas(cvs_wrap, bg='#FFFFFF', highlightthickness=1,
                        highlightbackground=C['border'],
                        width=cv_w, height=cv_h,
                        scrollregion=(0, 0, cv_w, cv_h))
        hb = tk.Scrollbar(cvs_wrap, orient='horizontal', command=cvs.xview)
        vb = tk.Scrollbar(cvs_wrap, orient='vertical', command=cvs.yview)
        cvs.configure(xscrollcommand=hb.set, yscrollcommand=vb.set)
        hb.pack(side='bottom', fill='x')
        vb.pack(side='right', fill='y')
        cvs.pack(side='left', fill='both', expand=True)

        # ─── Panel inferior con descripción ───
        bottom = tk.Frame(parent, bg=C['surface'], highlightthickness=1,
                          highlightbackground=C['border'])
        bottom.pack(fill=tk.X, padx=20, pady=(0, 14))
        lbl_sub = tk.Label(bottom, text="", font=font(10, italic=True),
                           bg=C['surface'], fg=C['text_l'], anchor='w')
        lbl_sub.pack(fill=tk.X, padx=14, pady=(10, 4))
        lbl_desc = tk.Label(bottom, text="", font=font_m(9),
                            bg=C['surface'], fg=C['text'],
                            justify='left', anchor='w', wraplength=1200)
                
        lbl_desc.pack(fill=tk.X, padx=14, pady=(0, 12))
        # ---- Mini-tabla de contribuciones del paso actual ----
        mini_frame = tk.Frame(bottom, bg=C['surface'])
        mini_frame.pack(fill=tk.X, padx=14, pady=(0, 10))

        lbl_mini = tk.Label(mini_frame, text="", font=font_m(8),
                            bg=C['surface'], fg=C['text_l'],
                            justify='left', anchor='w')
        lbl_mini.pack(fill=tk.X)


        # ─── Función de actualización ───
        def _actualizar():
            i = estado['idx']
            p = pasos[i]
            # Título
            lbl_titulo.config(text=p['titulo'])
            # Subtítulo y descripción
            lbl_sub.config(text=p['subtitulo'])
            lbl_desc.config(text=p['descripcion'])
            # Indicador
            lbl_ind.config(text=f"Paso {i+1} / {len(pasos)}")
            # Botones
            btn_prev.config(state=tk.NORMAL if i > 0 else tk.DISABLED,
                            bg=C['primary'] if i > 0 else C['surface'],
                            fg='white' if i > 0 else C['text_l'])
            btn_next.config(state=tk.NORMAL if i < len(pasos)-1 else tk.DISABLED,
                            bg=C['primary'] if i < len(pasos)-1 else C['surface'],
                            fg='white' if i < len(pasos)-1 else C['text_l'])
                        # Mini-tabla de contribuciones para el paso actual
            if p['P'] and sum(p['P']) > 0:
                contrib_g, contrib_l, _ = core_global.descomponer_flujos(
                    np.array(p['P']))
                lineas_txt = []
                for k in range(len(r['lines'])):
                    F = p['flows'][k]
                    # Contribuciones destacadas (las 3 mayores en valor absoluto)
                    contribs = []
                    for g in range(len(r['generators'])):
                        c = contrib_g[k][g]
                        if abs(c) > 0.1:
                            contribs.append(f"G{g+1}={c:+.1f}")
                    for i in range(r['n']):
                        c = contrib_l[k][i]
                        if abs(c) > 0.1:
                            contribs.append(f"D{i+1}={c:+.1f}")
                    txt = "  ".join(contribs) if contribs else "—"
                    lineas_txt.append(
                        f"L{k+1} (F={F:+.1f} MW):  {txt}")
                lbl_mini.config(text="\n".join(lineas_txt))
            else:
                lbl_mini.config(text="")
            # Redibujar canvas
            self._dibujar_unifilar_despacho(
                cvs, r['n'], r['generators'], r['loads'], r['lines'],
                p['P'], p['flows'], p['costo'], p['titulo'],
                cv_w, cv_h,
                con_redespacho=p['con_redespacho'],
                P_base=p['P_base'],
                precios=p['precios'])

        def _prev():
            if estado['idx'] > 0:
                estado['idx'] -= 1
                _actualizar()

        def _next():
            if estado['idx'] < len(pasos) - 1:
                estado['idx'] += 1
                _actualizar()

        btn_prev.config(command=_prev)
        btn_next.config(command=_next)

        # Atajos con flechas del teclado
        parent.bind_all('<Left>', lambda e: _prev())
        parent.bind_all('<Right>', lambda e: _next())

        # Pintar primer paso
        parent.after(50, _actualizar)

    # =================================================================
    # PESTAÑA CONTRIBUCIONES (desglose tipo slide 24-26)
    # =================================================================
    def _tab_contribuciones(self, parent, r):
        C = self.C
        n = r['n']
        gens = r['generators']
        lines = r['lines']
        P_fin = r['P_final']
        flows_fin = r['flows_final']
        contrib_gen = r.get('contrib_gen', [])
        contrib_load = r.get('contrib_load', [])

        # ---- Título ----
        tk.Label(parent,
                 text="F_ℓ = Σ_g (ISF_ℓ,bus_g · P_g) − Σ_i (ISF_ℓ,i · L_i)",
                 font=font_m(10), bg=C['bg'], fg=C['text_l']
                 ).pack(pady=(0, 10))
        tk.Label(parent,
                 text="F_ℓ = Σ_g (ISF_ℓ,bus_g · P_g) − Σ_i (ISF_ℓ,i · L_i)",
                 font=font_m(10), bg=C['bg'], fg=C['text_l']
                 ).pack(pady=(0, 10))

        # ---- Scroll ----
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas.find_all()[0],
                                                 width=e.width))
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        # ---- Bloque 1: contribución de cada generador a cada línea ----
        self._tabla(inner,
            "① CONTRIBUCIÓN DE CADA GENERADOR AL FLUJO (MW)",
            ["Línea"] + [f"G{i+1} (bus {g['bus']+1}, {P_fin[i]:.0f} MW)"
                         for i, g in enumerate(gens)] + ["Σ gen"],
            [[f"L{k+1} ({L['a']+1}→{L['b']+1})"]
             + [f"{contrib_gen[k][g]:+.2f}" for g in range(len(gens))]
             + [f"{sum(contrib_gen[k]):+.2f}"]
             for k, L in enumerate(lines)],
            color_col=None, ancho=14)

        # ---- Bloque 2: contribución de cada carga (retiro) ----
        self._tabla(inner,
            "② CONTRIBUCIÓN DE CADA CARGA AL FLUJO (MW)",
            ["Línea"] + [f"D{i+1} (bus {i+1}, {r['loads'][i]:.0f} MW)"
                         for i in range(n)] + ["Σ carga"],
            [[f"L{k+1} ({L['a']+1}→{L['b']+1})"]
             + [f"{contrib_load[k][i]:+.2f}" for i in range(n)]
             + [f"{sum(contrib_load[k]):+.2f}"]
             for k, L in enumerate(lines)],
            color_col=None, ancho=14)

        # ---- Bloque 3: suma = flujo total ----
        filas_totales = []
        for k, L in enumerate(lines):
            suma_gen = sum(contrib_gen[k])
            suma_load = sum(contrib_load[k])
            total = suma_gen + suma_load
            filas_totales.append([
                f"L{k+1}",
                f"{suma_gen:+.2f}",
                f"{suma_load:+.2f}",
                f"{total:+.2f}",
                f"{flows_fin[k]:+.2f}",
                "✓" if abs(total - flows_fin[k]) < 0.05 else "⚠",
            ])
        self._tabla(inner,
            "③ VERIFICACIÓN: Σ contribuciones = Flujo total (MW)",
            ["Línea", "Σ generadores", "Σ cargas",
             "Σ total", "Flujo DC", "✓"],
            filas_totales,
            color_col=3, ancho=15)

        # ---- Bloque 4: ejemplo didáctico (línea más cargada) ----
        # Encontrar la línea con mayor |flujo|
        k_max = max(range(len(lines)), key=lambda k: abs(flows_fin[k]))
        L = lines[k_max]
        F = flows_fin[k_max]

        ejemplo = tk.Frame(inner, bg='#F3E8FF', highlightthickness=1,
                            highlightbackground=C['purple'])
        ejemplo.pack(fill=tk.X, padx=20, pady=(18, 10))
        tk.Label(ejemplo,
                 text=f"EJEMPLO — DESCOMPOSICIÓN DE LA LÍNEA L{k_max+1} ({L['a']+1}→{L['b']+1})",
                 font=font(10, "bold"), bg='#F3E8FF', fg=C['purple'],
                 anchor='w').pack(fill=tk.X, padx=14, pady=(8, 4))
        tk.Label(ejemplo,
                 text=f"F_{k_max+1} = {F:+.2f} MW   →   se descompone así:",
                 font=font_m(10), bg='#F3E8FF', fg=C['text'],
                 anchor='w').pack(fill=tk.X, padx=14)

        for g in range(len(gens)):
            term = contrib_gen[k_max][g]
            if abs(term) < 1e-3:
                continue
            tk.Label(ejemplo,
                     text=f"    +  (ISF_{k_max+1},{gens[g]['bus']+1} · P_G{g+1})  "
                          f"=  {term:+.2f} MW",
                     font=font_m(9), bg='#F3E8FF', fg=C['text'],
                     anchor='w').pack(fill=tk.X, padx=28)
        for i in range(n):
            term = contrib_load[k_max][i]
            if abs(term) < 1e-3:
                continue
            tk.Label(ejemplo,
                     text=f"    −  (ISF_{k_max+1},{i+1} · L_D{i+1})  "
                          f"=  {term:+.2f} MW",
                     font=font_m(9), bg='#F3E8FF', fg=C['pq'],
                     anchor='w').pack(fill=tk.X, padx=28)
        tk.Label(ejemplo,
                 text=f"    ────────────────────────\n"
                      f"    F_{k_max+1}  =  {F:+.2f} MW",
                 font=font_m(10, "bold"), bg='#F3E8FF', fg=C['purple'],
                 anchor='w', justify='left').pack(fill=tk.X,
                                                   padx=28, pady=(0, 10))

        # ---- Nota final ----
        nota = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        nota.pack(fill=tk.X, padx=20, pady=(0, 20))
        tk.Label(nota,
                 text="¿Cómo se lee esto?\n"
                      "  • ISF_ℓ,i = factor de sensibilidad: cuánto cambia el "
                      "flujo de la línea ℓ si inyecto 1 MW en el nodo i "
                      "(y lo retiro del slack).\n"
                      "  • Cada generador g contribuye con "
                      "ISF_ℓ,bus_g · P_g MW al flujo de la línea ℓ.\n"
                      "  • Cada carga i sustrae ISF_ℓ,i · L_i MW.\n"
                      "  • La suma de todas las contribuciones = flujo total "
                      "F_ℓ calculado por el DC.\n"
                      "  • Si un término es positivo, empuja el flujo en el "
                      "sentido a→b. Si es negativo, empuja en b→a.",
                 font=font_m(9), bg=C['surface'], fg=C['text_l'],
                 justify='left', anchor='w').pack(fill=tk.X,
                                                   padx=14, pady=12)

    def _tab_esquematico(self, parent, r):
        """Reutiliza el mismo dibujo pero sin selector."""
        # Compatibilidad con la pestaña anterior
        self._tab_redespachos(parent, r)

    def _tab_resumen(self, parent, r):
        C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        info = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                        highlightbackground=C['border'])
        info.pack(fill=tk.X, padx=20, pady=(16, 12))
        estado = "CON CONGESTIÓN" if r['congestionadas'] else "ÓPTIMO SIN RESTRICCIONES"
        tk.Label(info,
                 text=f"Estado: {estado}   ·   Iteraciones de redespacho: {r['iteraciones']}"
                      f"   ·   Líneas congestionadas: {len(r['congestionadas'])}",
                 font=font_m(10), bg=C['surface'], fg=C['text'],
                 padx=14, pady=10, anchor='w').pack(fill=tk.X)

        kpis = tk.Frame(inner, bg=C['bg'])
        kpis.pack(fill=tk.X, padx=20, pady=(0, 14))

        def _card(t, v, s, col):
            card = tk.Frame(kpis, bg=C['panel'], highlightthickness=1,
                            highlightbackground=C['border'])
            card.pack(side=tk.LEFT, padx=6, fill=tk.Y)
            tk.Frame(card, bg=col, width=5).pack(side=tk.LEFT, fill=tk.Y)
            body = tk.Frame(card, bg=C['panel'])
            body.pack(side=tk.LEFT, padx=14, pady=10)
            tk.Label(body, text=t.upper(), font=font(8, "bold"),
                     bg=C['panel'], fg=C['text_l']).pack(anchor='w')
            tk.Label(body, text=v, font=font(12, "bold"),
                     bg=C['panel'], fg=col).pack(anchor='w', pady=(4, 0))
            tk.Label(body, text=s, font=font(8, italic=True),
                     bg=C['panel'], fg=C['text_l']).pack(anchor='w')

        _card("Costo ED (sin restr.)",
              f"${r['costo_ED']:.2f}/h", "Merit order puro", C['slack'])
        _card("Costo final",
              f"${r['costo_final']:.2f}/h", "Con redespacho", C['purple'])
        _card("Costo de seguridad",
              f"${r['costo_seguridad']:.2f}/h",
              "Cs = C_fin − C_ED", C['warning'])
        _card("Carga total",
              f"{sum(r['loads']):.1f} MW", "Suma de demandas", C['pq'])

        # Tabla de despacho
        self._tabla(inner, "DESPACHO FINAL POR GENERADOR",
            ["Gen #", "Nodo", "P (MW)", "Cap (MW)", "MC ($/MWh)",
             "P_ED (MW)", "Δ redespacho (MW)"],
            [[f"G{i+1}", g['bus']+1, f"{r['P_final'][i]:.2f}",
              f"{g['cap']:.1f}", f"{g['mc']:.2f}",
              f"{r['P_ED'][i]:.2f}",
              f"{r['P_final'][i] - r['P_ED'][i]:+.2f}"]
             for i, g in enumerate(r['generators'])],
            color_col=2)

    def _tab_despacho(self, parent, r):
        C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        self._tabla(inner, "DESPACHO SIN RESTRICCIONES (MERIT ORDER)",
            ["Gen", "Nodo", "P (MW)", "MC ($/MWh)", "Costo parcial ($/h)"],
            [[f"G{i+1}", g['bus']+1, f"{r['P_ED'][i]:.2f}",
              f"{g['mc']:.2f}",
              f"{r['P_ED'][i] * g['mc']:.2f}"]
             for i, g in enumerate(r['generators'])])
        self._tabla(inner, "DESPACHO FINAL (CON RESTRICCIONES)",
            ["Gen", "Nodo", "P (MW)", "MC ($/MWh)", "Costo parcial ($/h)"],
            [[f"G{i+1}", g['bus']+1, f"{r['P_final'][i]:.2f}",
              f"{g['mc']:.2f}",
              f"{r['P_final'][i] * g['mc']:.2f}"]
             for i, g in enumerate(r['generators'])],
            color_col=2)

    def _tab_flujos(self, parent, r):
        C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        rows = []
        for k, L in enumerate(r['lines']):
            f_ed = r['flows_ED'][k]
            f_fin = r['flows_final'][k]
            estado = "CONGESTIONADA" if abs(f_fin) >= L['cap'] - 1e-3 else "OK"
            rows.append([f"L{k+1}", f"{L['a']+1}→{L['b']+1}",
                         f"{L['x']:.3f}", f"{L['cap']:.1f}",
                         f"{f_ed:+.2f}", f"{f_fin:+.2f}", estado])
        self._tabla(inner, "FLUJOS POR LÍNEA (MW)",
            ["Línea", "Tramo", "x (p.u.)", "Cap (MW)",
             "F_ED", "F_final", "Estado"],
            rows, color_col=5)

    def _tab_precios(self, parent, r):
        C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        self._tabla(inner, "PRECIOS MARGINALES NODALES",
            ["Nodo", "Carga (MW)", "π ($/MWh)"],
            [[str(i+1), f"{r['loads'][i]:.1f}",
              f"{r['precios_nodales'][i]:.2f}"]
             for i in range(r['n'])], color_col=2)

    def _tab_isf(self, parent, r):
        C = self.C
        canvas = tk.Canvas(parent, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        headers = ["Línea"] + [f"N{i+1}" for i in range(r['n'])]
        rows = []
        for k in range(len(r['lines'])):
            rows.append([f"L{k+1}"] +
                        [f"{r['ISF'][k][i]:+.4f}" for i in range(r['n'])])
        self._tabla(inner, "INJECTION SHIFT FACTORS (ISF)",
                    headers, rows)

    def _tabla(self, parent, titulo, headers, rows, color_col=None, ancho=14):
        C = self.C
        f_titulo = tk.Frame(parent, bg=C['surface'], highlightthickness=1,
                            highlightbackground=C['border'])
        f_titulo.pack(fill=tk.X, padx=20, pady=(10, 8))
        tk.Frame(f_titulo, bg=C['purple'], width=5).pack(side=tk.LEFT, fill=tk.Y)
        tk.Label(f_titulo, text=titulo, font=font(10, "bold"),
                 bg=C['surface'], fg=C['text'], pady=8
                 ).pack(side=tk.LEFT, padx=12)

        tabla = tk.Frame(parent, bg=C['border'])
        tabla.pack(fill=tk.X, padx=20, pady=(0, 14))
        head = tk.Frame(tabla, bg=C['primary_d'])
        head.pack(fill=tk.X)
        for h in headers:
            tk.Label(head, text=h, font=font(9, "bold"),
                     bg=C['primary_d'], fg='white', width=ancho,
                     pady=8).pack(side=tk.LEFT)
        body = tk.Frame(tabla, bg=C['border'])
        body.pack(fill=tk.X)
        for i, fila in enumerate(rows):
            bg_r = C['row_even'] if i % 2 == 0 else C['row_odd']
            f = tk.Frame(body, bg=C['border'])
            f.pack(fill=tk.X, pady=1)
            for c, val in enumerate(fila):
                fg_c = C['purple'] if (color_col is not None and c == color_col) \
                       else C['text']
                tk.Label(f, text=str(val), font=font_m(9),
                         bg=bg_r, fg=fg_c, width=ancho, pady=6
                         ).pack(side=tk.LEFT)

    # =================================================================
    # TEORÍA
    # =================================================================
    def mostrar_teoria(self):
        C = self.C
        win = tk.Toplevel(self.root)
        win.title("Teoría — Despacho Económico")
        win.geometry("950x700")
        win.configure(bg=C['bg'])
        hdr = tk.Frame(win, bg=C['primary_d'], height=64)
        hdr.pack(fill=tk.X); hdr.pack_propagate(False)
        tk.Label(hdr, text="  TEORÍA — DESPACHO ECONÓMICO CON CONGESTIÓN",
                 font=font(13, "bold"), bg=C['primary_d'], fg='white'
                 ).pack(side=tk.LEFT, padx=20, pady=18)

        canvas = tk.Canvas(win, bg=C['bg'], highlightthickness=0)
        sb = tk.Scrollbar(win, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=C['bg'])
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        def _sec(t, txt):
            f = tk.Frame(inner, bg=C['surface'], highlightthickness=1,
                         highlightbackground=C['border'])
            f.pack(fill=tk.X, padx=20, pady=8)
            tk.Frame(f, bg=C['purple'], width=5).pack(side=tk.LEFT, fill=tk.Y)
            body = tk.Frame(f, bg=C['surface'])
            body.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=14, pady=10)
            tk.Label(body, text=t, font=font(11, "bold"),
                     bg=C['surface'], fg=C['primary'],
                     anchor='w').pack(fill=tk.X)
            tk.Label(body, text=txt, font=font_m(10),
                     bg=C['surface'], fg=C['text'], justify='left',
                     anchor='w', wraplength=850).pack(fill=tk.X, pady=(4, 0))

        _sec("1. DESPACHO SIN RESTRICCIONES",
             "Ordenar generadores por costo marginal ascendente (merit order). "
             "Despachar en ese orden hasta cubrir la demanda total. Este es "
             "el óptimo económico sin considerar la red.")
        _sec("2. FLUJO DC",
             "θ = B_red⁻¹ · P_iny\n"
             "F_km = (θ_k − θ_m) / x_km\n"
             "Permite verificar si algún flujo excede la capacidad.")
        _sec("3. INJECTION SHIFT FACTOR (ISF)",
             "ISF_ℓ,i = (X[a][i] − X[b][i]) / x_ℓ\n"
             "Indica cuánto cambia el flujo en la línea ℓ si inyecto 1 MW "
             "en el nodo i y lo extraigo del slack.")
        _sec("4. CONGESTIÓN Y REDESPACHO",
             "Si |F_ℓ| > cap_ℓ → hay congestión. Se busca el par (u, d) de "
             "generadores (uno que sube, otro que baja) que reduzca el flujo "
             "en la línea sobrecargada al menor costo. Se itera hasta que "
             "todas las líneas cumplan.")
        _sec("5. COSTO DE SEGURIDAD",
             "Cs = C_restringido − C_ED\n"
             "Es el sobrecosto de operar la red de forma segura. "
             "En el ejemplo de las slides, Cs = $187.50/h.")
        _sec("6. PRECIOS NODALES",
             "π_k = costo marginal de servir 1 MW adicional en el nodo k, "
             "respetando las restricciones activas.\n"
             "Sin congestión: π_k = MC del generador marginal (uniforme).\n"
             "Con congestión: los precios difieren por nodo. En el ejemplo: "
             "π_1 = $7.50, π_2 = $11.25, π_3 = $10.00 por MWh.")

        tk.Button(inner, text="Cerrar", command=win.destroy,
                  bg=C['primary'], fg='white', font=font(10, "bold"),
                  relief=tk.FLAT, padx=24, pady=8, cursor='hand2'
                  ).pack(pady=20)

    # =================================================================
    # GUARDAR / ABRIR / EXPORTAR / NAVEGACIÓN
    # =================================================================
    def _nuevo_proyecto(self):
        if messagebox.askyesno("Nuevo proyecto",
            "¿Crear un nuevo proyecto?\n\nSe perderán los datos actuales."):
            self.generar_tablas()
            self._set_status("Nuevo proyecto creado.", 'info')

    def _guardar_proyecto(self):
        import json
        path = filedialog.asksaveasfilename(
            defaultextension=".vatia_despacho.json",
            filetypes=[("Proyecto Despacho (VATIA)", "*.vatia_despacho.json"),
                       ("JSON", "*.json")])
        if not path: return
        try:
            n = int(self.entry_n.get())
            data = {
                'version': 1, 
                'modo': 'despacho', 
                'n': n,
                'n_gen': self.n_gen,
                'loads': [w.get() for w in self.load_widgets],
                'lines': [[e.get() for e in row] for row in self.line_widgets],
                'gens': [[e.get() for e in row] for row in self.gen_widgets],
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            self._set_status(f"Guardado: {os.path.basename(path)}", 'ok')
        except Exception as e:
            messagebox.showerror("Error al guardar", str(e))

    def _abrir_proyecto(self):
        import json
        path = filedialog.askopenfilename(
            filetypes=[("Proyecto Despacho (VATIA)", "*.vatia_despacho.json"),
                       ("JSON", "*.json")])
        if not path: return
        try:
            with open(path, encoding='utf-8') as f:
                data = json.load(f)

            # --- Sincronizar ambos campos ANTES de generar_tablas ---
            self.entry_n.delete(0, tk.END)
            self.entry_n.insert(0, str(data['n']))

            n_gen_guardado = len(data.get('gens', []))
            self.entry_n_gen.delete(0, tk.END)
            self.entry_n_gen.insert(0, str(n_gen_guardado or 3))

            self.generar_tablas()

            # ... el resto igual (cargar loads, lines, gens)
            for i, v in enumerate(data.get('loads', [])):
                if i < len(self.load_widgets):
                    self.load_widgets[i].delete(0, tk.END)
                    self.load_widgets[i].insert(0, v)
            for i, fila in enumerate(data.get('lines', [])):
                if i < len(self.line_widgets):
                    for j, v in enumerate(fila):
                        if j < len(self.line_widgets[i]):
                            self.line_widgets[i][j].delete(0, tk.END)
                            self.line_widgets[i][j].insert(0, v)
            for i, fila in enumerate(data.get('gens', [])):
                if i < len(self.gen_widgets):
                    for j, v in enumerate(fila):
                        if j < len(self.gen_widgets[i]):
                            self.gen_widgets[i][j].delete(0, tk.END)
                            self.gen_widgets[i][j].insert(0, v)
            self._update_sidebar()
            self._dibujar_preview()
            self._set_status(f"Cargado: {os.path.basename(path)}", 'ok')
        except Exception as e:
            import traceback; traceback.print_exc()
            messagebox.showerror("Error al abrir", str(e))

    def _exportar_excel(self):
        if not self.resultados:
            messagebox.showinfo("Info", "Calcula primero el despacho.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")])
        if not path: return
        try:
            r = self.resultados
            wb = openpyxl.Workbook()
            ws = wb.active; ws.title = "Despacho"
            headers = ["Gen", "Nodo", "P_ED (MW)", "P_final (MW)",
                       "Cap (MW)", "MC ($/MWh)", "Costo ($/h)"]
            for j, h in enumerate(headers, 1):
                ws.cell(row=1, column=j, value=h)
            for i, g in enumerate(r['generators']):
                ws.cell(row=i+2, column=1, value=f"G{i+1}")
                ws.cell(row=i+2, column=2, value=g['bus']+1)
                ws.cell(row=i+2, column=3, value=r['P_ED'][i])
                ws.cell(row=i+2, column=4, value=r['P_final'][i])
                ws.cell(row=i+2, column=5, value=g['cap'])
                ws.cell(row=i+2, column=6, value=g['mc'])
                ws.cell(row=i+2, column=7,
                        value=r['P_final'][i] * g['mc'])
            ws2 = wb.create_sheet("Flujos")
            for k, L in enumerate(r['lines']):
                ws2.cell(row=k+2, column=1, value=f"L{k+1}")
                ws2.cell(row=k+2, column=2, value=L['a']+1)
                ws2.cell(row=k+2, column=3, value=L['b']+1)
                ws2.cell(row=k+2, column=4, value=L['x'])
                ws2.cell(row=k+2, column=5, value=L['cap'])
                ws2.cell(row=k+2, column=6, value=r['flows_final'][k])
            ws3 = wb.create_sheet("Precios")
            for i in range(r['n']):
                ws3.cell(row=i+2, column=1, value=i+1)
                ws3.cell(row=i+2, column=2, value=r['loads'][i])
                ws3.cell(row=i+2, column=3, value=r['precios_nodales'][i])
            wb.save(path)
            self._set_status("Excel exportado.", 'ok')
            if messagebox.askyesno("Exportado",
                f"Guardado en:\n{path}\n\n¿Abrir?"):
                try: os.startfile(path)
                except Exception: pass
        except Exception as e:
            import traceback; traceback.print_exc()
            messagebox.showerror("Error", f"No se pudo exportar:\n{e}")

    def _volver_al_menu(self):
        if not messagebox.askyesno("Confirmar",
            "¿Volver al menú principal?\nSe perderán los datos no guardados."):
            return
        try:
            if self._sidebar_after_id:
                self.root.after_cancel(self._sidebar_after_id)
                self._sidebar_after_id = None
        except Exception:
            pass
        for w in self.root.winfo_children():
            w.destroy()
        try:
            def abrir_simulador(root):
                root.title("Vatia — Flujos de Potencia (Newton-Raphson)")
                root.geometry("1500x900")
                root.minsize(1180, 720)
                root.configure(bg="#EDF1F6")
                app = AppFlujoPotencia(root)
                root.app = app
            DashboardScreen(self.root, abrir_simulador)
        except Exception:
            import traceback; traceback.print_exc()

    def _on_close(self):
        try:
            if self._sidebar_after_id:
                self.root.after_cancel(self._sidebar_after_id)
                self._sidebar_after_id = None
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass
    

# =====================================================================
# SPLASH SCREEN
# =====================================================================
class SplashScreen:
    BG = '#F7F9FC'; BG2 = '#EDF1F6'
    GRID = '#E5E7EB'; GRID_HL = '#CBD5E1'
    CYAN = '#1E3A5F'; CYAN_DIM = '#475569'; CYAN_DK = '#94A3B8'
    AMBER = '#0E7490'; WHITE = '#FFFFFF'
    TEXT = '#1F2937'; TEXT_DIM = '#64748B'
    SLACK = '#047857'; PV = '#6D28D9'; PQ = '#C2410C'
    WIRE = '#94A3B8'; PARTICLE = '#0E7490'; FLASH = '#FFFFFF'

    def __init__(self, splash_root, on_finish, duration_frames=300, frame_ms=18):
        self.root = splash_root
        self.on_finish = on_finish
        self.duration = duration_frames
        self.frame_ms = frame_ms
        self.tick = 0
        self.particles = []
        self.arcs = []
        self.rings = []
        self.scan_y = 0
        self.rotation = 0.0
        self.finishing = False
        self.W, self.H = 980, 640

        try: self.root.overrideredirect(True)
        except Exception: pass
        try: self.root.attributes('-topmost', True)
        except Exception: pass

        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - self.W) // 2
        y = (sh - self.H) // 2
        self.root.geometry(f"{self.W}x{self.H}+{x}+{y}")
        self.root.configure(bg=self.BG)

        self.root.update_idletasks()
        self.root.lift()
        try: self.root.focus_force()
        except Exception: pass
        self.root.update()

        outer = tk.Frame(self.root, bg=self.CYAN)
        outer.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        self.cv = tk.Canvas(outer, width=self.W-4, height=self.H-4,
                            bg=self.BG, highlightthickness=0)
        self.cv.pack(fill=tk.BOTH, expand=True)

        self._build_background()
        self._build_network()
        self._build_typography()
        self._build_progress()
        self.root.after(self.frame_ms, self._loop)

    def _build_background(self):
        W, H = self.W - 4, self.H - 4
        self.cw, self.ch = W, H
        for i in range(60):
            t = i / 59
            color = interp_color(self.BG, self.BG2, t)
            self.cv.create_rectangle(0, i*H/60, W, (i+1)*H/60+1,
                                       fill=color, outline='')
        for x in range(0, W, 32):
            self.cv.create_line(x, 0, x, H, fill=self.GRID, width=1)
        for y in range(0, H, 32):
            self.cv.create_line(0, y, W, y, fill=self.GRID, width=1)
        for x in range(0, W, 160):
            self.cv.create_line(x, 0, x, H, fill=self.GRID_HL, width=1)
        for y in range(0, H, 160):
            self.cv.create_line(0, y, W, y, fill=self.GRID_HL, width=1)
        self._circuit_corner(20, 20, 1, 1)
        self._circuit_corner(W-20, 20, -1, 1)
        self._circuit_corner(20, H-20, 1, -1)
        self._circuit_corner(W-20, H-20, -1, -1)
        c = 28
        for (cx, cy, sx, sy) in [(10,10,1,1), (W-10,10,-1,1),
                                   (10,H-10,1,-1), (W-10,H-10,-1,-1)]:
            self.cv.create_line(cx, cy, cx+c*sx, cy, fill=self.CYAN, width=2)
            self.cv.create_line(cx, cy, cx, cy+c*sy, fill=self.CYAN, width=2)

    def _circuit_corner(self, x0, y0, sx, sy):
        segs = [[(0,0),(30,0),(30,20),(60,20)],
                [(0,30),(20,30),(20,50)],
                [(50,0),(50,15),(70,15),(70,30)]]
        for seg in segs:
            pts = []
            for (px, py) in seg:
                pts.extend([x0+px*sx, y0+py*sy])
            self.cv.create_line(*pts, fill=self.CYAN_DK, width=1,
                                capstyle=tk.ROUND)

    def _build_network(self):
        W, H = self.cw, self.ch
        self.cx = W/2
        self.cy = H/2 + 20
        self.R = 150
        self.n_nodos = 6
        self.nodos = [(0, 0)] * self.n_nodos
        self._update_node_positions(0.0)
        self.edges = []
        for i in range(self.n_nodos):
            self.edges.append((i, (i+1) % self.n_nodos))
        self.edges += [(0,2), (2,4), (4,0)]
        self.nodo_col = [self.SLACK, self.PV, self.PQ, self.PQ, self.PV, self.SLACK]

        self.ring_items = []
        for k in range(4):
            r = 40 + k*35
            it = self.cv.create_oval(self.cx-r, self.cy-r, self.cx+r, self.cy+r,
                                       outline=self.BG, width=1.5)
            self.ring_items.append(it)

        self.line_items = []
        for _ in self.edges:
            it = self.cv.create_line(0, 0, 0, 0, fill=self.BG,
                                       width=2.5, capstyle=tk.ROUND)
            self.line_items.append(it)

        self.node_halo_items = []
        self.node_items = []
        for i in range(self.n_nodos):
            halo = self.cv.create_oval(0, 0, 0, 0, fill=self.BG, outline='')
            self.node_halo_items.append(halo)
            it = self.cv.create_oval(0, 0, 0, 0, fill=self.BG,
                                       outline=self.BG, width=2.5)
            self.node_items.append(it)

        self.core_layers = []
        for r in (90, 70, 50, 34, 22):
            it = self.cv.create_oval(self.cx-r, self.cy-r, self.cx+r, self.cy+r,
                                       fill=self.BG, outline='')
            self.core_layers.append((it, r))

        self.logo_item = self.cv.create_text(self.cx, self.cy, text="V",
                                               font=('Segoe UI', 60, 'bold'),
                                               fill=self.BG)

        self.tech_ring_items = []
        for k, r in enumerate((105, 118)):
            arc = self.cv.create_arc(self.cx-r, self.cy-r, self.cx+r, self.cy+r,
                                       start=0, extent=60, style=tk.ARC,
                                       outline=self.BG, width=2)
            self.tech_ring_items.append((arc, r, k))
            arc2 = self.cv.create_arc(self.cx-r, self.cy-r, self.cx+r, self.cy+r,
                                        start=180, extent=40, style=tk.ARC,
                                        outline=self.BG, width=2)
            self.tech_ring_items.append((arc2, r, k))

    def _build_typography(self):
        W = self.cw
        self.title_item = self.cv.create_text(W/2, 68, text="V A T I A",
                                                font=('Segoe UI', 38, 'bold'),
                                                fill=self.BG)
        self.title_line_l = self.cv.create_line(W/2, 108, W/2, 108,
                                                  fill=self.CYAN, width=3)
        self.title_line_r = self.cv.create_line(W/2, 108, W/2, 108,
                                                  fill=self.CYAN, width=3)
        self.sub_item = self.cv.create_text(W/2, 130,
            text="SIMULADOR DE SISTEMAS DE POTENCIA",
            font=('Segoe UI', 11, 'bold'), fill=self.BG, anchor='center')
        self.sub_sub = self.cv.create_text(W/2, 152,
            text="Newton-Raphson  ·  Flujo de Carga  ·  Análisis Nodal",
            font=('Segoe UI', 9, 'italic'), fill=self.BG)
        self.author_item = self.cv.create_text(W-24, self.ch-24, text="por Kevin",
                                                 font=('Segoe UI', 10, 'italic'),
                                                 fill=self.BG, anchor='se')
        self.cv.create_text(24, self.ch-24, text="v1.0",
                             font=('Consolas', 9), fill=self.CYAN_DK, anchor='sw')
        self.cv.create_text(24, 22, text="SISTEMA ONLINE",
                             font=('Consolas', 9, 'bold'),
                             fill=self.CYAN_DK, anchor='nw')
        self.cv.create_text(W-24, 22, text="BUILD 2026.10",
                             font=('Consolas', 9), fill=self.CYAN_DK, anchor='ne')

    def _build_progress(self):
        W, H = self.cw, self.ch
        bar_w = 560
        self.bar_x0 = (W - bar_w) // 2
        self.bar_y0 = H - 78
        self.bar_w = bar_w
        self.bar_h = 4
        self.cv.create_rectangle(self.bar_x0-1, self.bar_y0-1,
                                   self.bar_x0+bar_w+1, self.bar_y0+self.bar_h+1,
                                   fill='#E5E7EB', outline='')
        self.bar_fill = self.cv.create_rectangle(self.bar_x0, self.bar_y0,
                                                   self.bar_x0,
                                                   self.bar_y0+self.bar_h,
                                                   fill=self.CYAN, outline='')
        self.bar_glow = self.cv.create_rectangle(self.bar_x0, self.bar_y0-1,
                                                   self.bar_x0,
                                                   self.bar_y0+self.bar_h+1,
                                                   fill=self.WHITE, outline='')
        self.status_item = self.cv.create_text(W/2, H-42, text="",
                                                 font=('Consolas', 10),
                                                 fill=self.CYAN)
        self.estados = [
            (10, "> Inicializando motor de cálculo..."),
            (55, "> Cargando topología de red..."),
            (100, "> Construyendo matriz Ybus..."),
            (150, "> Ensamblando Jacobiano..."),
            (200, "> Ejecutando iteración Newton-Raphson..."),
            (250, "> Verificando convergencia..."),
            (285, "> Sistema listo. Iniciando simulador.")]

    def _update_node_positions(self, rot):
        for i in range(self.n_nodos):
            ang = 2*math.pi*i/self.n_nodos - math.pi/2 + rot
            self.nodos[i] = (self.cx + self.R*math.cos(ang),
                              self.cy + self.R*math.sin(ang))

    def _spawn_arc(self, a, b):
        x1, y1 = self.nodos[a]
        x2, y2 = self.nodos[b]
        pts = []
        seg = 7
        for k in range(seg+1):
            t = k/seg
            px = x1 + (x2-x1)*t
            py = y1 + (y2-y1)*t
            if 0 < k < seg:
                perp_x = -(y2-y1)
                perp_y = (x2-x1)
                L = math.hypot(perp_x, perp_y) or 1.0
                perp_x /= L
                perp_y /= L
                off = (hash((self.tick, k, a, b)) % 12) - 6
                px += perp_x*off
                py += perp_y*off
            pts.extend([px, py])
        it = self.cv.create_line(*pts, fill=self.CYAN, width=2, smooth=True)
        self.arcs.append({'item': it, 'life': 4})

    def _loop(self):
        try:
            t = self.tick
            self.tick += 1
            if self.finishing:
                return
            if t >= 90:
                self.rotation += 0.012
                self._update_node_positions(self.rotation)
            self._anim_typography(t)
            self._anim_background_pulse(t)
            self._anim_network(t)
            self._anim_core(t)
            self._anim_arcs(t)
            self._anim_particles(t)
            self._anim_scanline(t)
            self._anim_progress(t)
            if t < self.duration:
                self.root.after(self.frame_ms, self._loop)
            else:
                self._fade_out()
        except Exception:
            import traceback
            traceback.print_exc()
            self._finish()

    def _anim_typography(self, t):
        if t < 25:
            a = t / 25
            self.cv.itemconfig(self.title_item,
                                fill=interp_color(self.BG, self.CYAN, a))
        if 15 < t < 45:
            a = (t-15)/30
            hw = 140*a
            self.cv.coords(self.title_line_l, self.cw/2-hw, 108, self.cw/2, 108)
            self.cv.coords(self.title_line_r, self.cw/2, 108, self.cw/2+hw, 108)
        elif t >= 45:
            self.cv.coords(self.title_line_l, self.cw/2-140, 108, self.cw/2, 108)
            self.cv.coords(self.title_line_r, self.cw/2, 108, self.cw/2+140, 108)
        if 30 < t < 55:
            a = (t-30)/25
            self.cv.itemconfig(self.sub_item,
                                fill=interp_color(self.BG, self.CYAN, a))
        if 40 < t < 70:
            a = (t-40)/30
            self.cv.itemconfig(self.sub_sub,
                                fill=interp_color(self.BG, self.TEXT_DIM, a))
        if t > self.duration - 60:
            a = min(1.0, (t - (self.duration - 60)) / 30)
            self.cv.itemconfig(self.author_item,
                                fill=interp_color(self.BG, self.AMBER, a))

    def _anim_background_pulse(self, t):
        if t > 90 and t % 55 == 0 and len(self.rings) < 4:
            r0 = 30
            it = self.cv.create_oval(self.cx-r0, self.cy-r0,
                                       self.cx+r0, self.cy+r0,
                                       outline=self.CYAN_DK, width=1.5)
            self.rings.append({'item': it, 'r': r0, 'life': 80})
        for rg in list(self.rings):
            rg['r'] += 3.2
            rg['life'] -= 1
            r = rg['r']
            self.cv.coords(rg['item'], self.cx-r, self.cy-r,
                           self.cx+r, self.cy+r)
            col_t = min(1.0, max(0.0, (rg['life']-20)/60))
            self.cv.itemconfig(rg['item'],
                                outline=interp_color(self.BG, self.CYAN_DK, col_t))
            if rg['life'] <= 0:
                try: self.cv.delete(rg['item'])
                except Exception: pass
                self.rings.remove(rg)
        for k, it in enumerate(self.ring_items):
            start = 80 + k*8
            if start <= t < start+20:
                a = (t-start)/20
                self.cv.itemconfig(it,
                                    outline=interp_color(self.BG, self.CYAN_DK, a*0.6))
            elif t >= start+20:
                self.cv.itemconfig(it, outline=self.CYAN_DK)

    def _anim_network(self, t):
        for i, it in enumerate(self.line_items):
            start = 50 + i*6
            if start <= t < start+15:
                a = (t-start)/15
                self.cv.itemconfig(it, fill=interp_color(self.BG, self.WIRE, a))
            elif t >= start+15:
                self.cv.itemconfig(it, fill=self.WIRE)
            a_i, b_i = self.edges[i]
            x1, y1 = self.nodos[a_i]
            x2, y2 = self.nodos[b_i]
            self.cv.coords(it, x1, y1, x2, y2)
        node_start_base = 50 + len(self.line_items)*6 + 10
        for i, it in enumerate(self.node_items):
            start = node_start_base + i*10
            x, y = self.nodos[i]
            if start <= t < start+20:
                a = (t-start)/20
                r = 18*a + 4*(1-a)
                self.cv.coords(it, x-r, y-r, x+r, y+r)
                self.cv.itemconfig(it,
                                    fill=interp_color(self.BG, self.nodo_col[i], a),
                                    outline=interp_color(self.BG, self.WHITE, a))
            elif t >= start+20:
                r = 18
                self.cv.coords(it, x-r, y-r, x+r, y+r)
                self.cv.itemconfig(it, fill=self.nodo_col[i], outline=self.WHITE)
                pulse = 0.20 + 0.15*abs(math.sin(t*0.06 + i))
                hr = 28 + 4*abs(math.sin(t*0.06 + i))
                self.cv.coords(self.node_halo_items[i], x-hr, y-hr, x+hr, y+hr)
                self.cv.itemconfig(self.node_halo_items[i],
                                    fill=interp_color(self.BG, self.nodo_col[i], pulse))

    def _anim_core(self, t):
        if t < 60: return
        a = min(1.0, (t-60)/60)
        pulse = 0.5 + 0.5*abs(math.sin(t*0.055))
        for i, (it, base_r) in enumerate(self.core_layers):
            layer_pulse = pulse if i % 2 == 0 else (1-pulse)
            r = base_r * (0.75 + 0.25*layer_pulse) * a
            col = interp_color(self.BG, self.CYAN, a * (0.15 + 0.10*(4-i)))
            self.cv.coords(it, self.cx-r, self.cy-r, self.cx+r, self.cy+r)
            self.cv.itemconfig(it, fill=col, outline='')
        logo_col = interp_color(self.BG, self.CYAN, a * (0.8 + 0.2*pulse))
        self.cv.itemconfig(self.logo_item, fill=logo_col)
        for (arc, r, k) in self.tech_ring_items:
            if t < 100: continue
            a2 = min(1.0, (t-100)/40)
            self.cv.itemconfig(arc,
                                outline=interp_color(self.BG, self.CYAN, a2*0.9))
            new_start = (t * (2 if k == 0 else -3) + (0 if k == 0 else 180)) % 360
            self.cv.itemconfig(arc, start=new_start,
                                extent=60 if k == 0 else 40)

    def _anim_arcs(self, t):
        if 120 < t < self.duration - 30 and t % 18 == 0:
            e = self.edges[(t*7) % len(self.edges)]
            self._spawn_arc(e[0], e[1])
        for arc in list(self.arcs):
            arc['life'] -= 1
            a = max(0.0, arc['life']/4)
            self.cv.itemconfig(arc['item'],
                                fill=interp_color(self.BG, self.CYAN, a))
            if arc['life'] <= 0:
                try: self.cv.delete(arc['item'])
                except Exception: pass
                self.arcs.remove(arc)

    def _anim_particles(self, t):
        if t < 130: return
        if t % 4 == 0 and len(self.particles) < 80:
            e_idx = (t//4) % len(self.edges)
            a, b = self.edges[e_idx]
            head = self.cv.create_oval(0,0,0,0, fill=self.CYAN, outline='')
            tail = self.cv.create_oval(0,0,0,0, fill=self.PARTICLE, outline='')
            self.particles.append({'head': head, 'tail': tail,
                                    'a': a, 'b': b, 't': 0.0,
                                    'speed': 0.018 + ((t*13) % 8)*0.003})
        for p in list(self.particles):
            p['t'] += p['speed']
            if p['t'] >= 1.0:
                try:
                    self.cv.delete(p['head'])
                    self.cv.delete(p['tail'])
                except Exception:
                    pass
                self.particles.remove(p)
                continue
            x1, y1 = self.nodos[p['a']]
            x2, y2 = self.nodos[p['b']]
            x = x1 + (x2-x1)*p['t']
            y = y1 + (y2-y1)*p['t']
            hr = 3.5
            self.cv.coords(p['head'], x-hr, y-hr, x+hr, y+hr)
            tt = max(0.0, p['t']-0.08)
            xt = x1 + (x2-x1)*tt
            yt = y1 + (y2-y1)*tt
            tr = 5.5
            self.cv.coords(p['tail'], xt-tr, yt-tr, xt+tr, yt+tr)

    def _anim_scanline(self, t):
        if t < 40: return
        if t % 90 < 60:
            progress = (t % 90)/60.0
            y = progress * self.ch
            if not hasattr(self, 'scan_item'):
                self.scan_item = self.cv.create_line(0, y, self.cw, y,
                                                       fill=self.CYAN_DK, width=1)
            else:
                self.cv.coords(self.scan_item, 0, y, self.cw, y)
                self.cv.itemconfig(self.scan_item, fill=self.CYAN_DK)
        else:
            if hasattr(self, 'scan_item'):
                try: self.cv.delete(self.scan_item)
                except Exception: pass
                del self.scan_item

    def _anim_progress(self, t):
        progress = min(1.0, t/self.duration)
        x1 = self.bar_x0
        x2 = self.bar_x0 + self.bar_w*progress
        self.cv.coords(self.bar_fill, x1, self.bar_y0, x2,
                       self.bar_y0+self.bar_h)
        self.cv.coords(self.bar_glow, x2-3, self.bar_y0-2, x2+3,
                       self.bar_y0+self.bar_h+2)
        txt_full = ""
        for (frame, s) in self.estados:
            if t >= frame: txt_full = s
        if txt_full:
            chars = min(len(txt_full), max(0, (t-10)*2))
            self.cv.itemconfig(self.status_item, text=txt_full[:chars])
            self.cv.itemconfig(self.status_item,
                                fill=self.CYAN if (t//10) % 2 == 0
                                     else self.CYAN_DIM)

    def _fade_out(self):
        flash = tk.Frame(self.root, bg=self.WHITE)
        flash.place(x=0, y=0, relwidth=1, relheight=1)
        def flash_step(step=0):
            if step >= 6:
                try: flash.destroy()
                except Exception: pass
                smooth_fade()
                return
            alpha = 1.0 - step/6
            try:
                flash.configure(bg=interp_color(self.BG, self.WHITE, alpha))
            except Exception:
                pass
            self.root.after(30, lambda: flash_step(step+1))
        def smooth_fade(alpha=1.0):
            alpha -= 0.08
            if alpha <= 0:
                self._finish()
                return
            try: self.root.attributes('-alpha', alpha)
            except Exception: pass
            self.root.after(20, lambda: smooth_fade(alpha))
        flash_step()

    def _finish(self):
        if self.finishing: return
        self.finishing = True
        try: self.root.destroy()
        except Exception: pass
        self.on_finish()


# =====================================================================
# ARRANQUE
# =====================================================================
if __name__ == "__main__":
    import traceback

    def launch_dashboard(splash_root):
        """Crea el dashboard como nueva root y aplica icono."""
        try:
            dash_root = tk.Tk()

            def abrir_simulador(root):
                try:
                    root.title("Vatia — Flujos de Potencia (Newton-Raphson)")
                    root.geometry("1500x900")
                    root.minsize(1180, 720)
                    root.configure(bg="#EDF1F6")
                    app = AppFlujoPotencia(root)
                    root.app = app
                except Exception:
                    traceback.print_exc()
                    messagebox.showerror(
                        "Error",
                        f"Error al abrir el módulo:\n\n{traceback.format_exc()}")

            dash = DashboardScreen(dash_root, abrir_simulador)

            # Icono para el dashboard
            try:
                from PIL import Image, ImageDraw, ImageTk
                ico_path = os.path.join(tempfile.gettempdir(), "vatia_icon.ico")
                if not os.path.exists(ico_path):
                    img = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
                    draw = ImageDraw.Draw(img)
                    pts = []
                    for i in range(6):
                        ang = math.pi * i / 3 - math.pi / 2
                        pts.append((128 + 115 * math.cos(ang),
                                    128 + 115 * math.sin(ang)))
                    draw.polygon(pts, fill=(30, 58, 95, 255))
                    pts2 = []
                    for i in range(6):
                        ang = math.pi * i / 3 - math.pi / 2
                        pts2.append((128 + 98 * math.cos(ang),
                                     128 + 98 * math.sin(ang)))
                    draw.polygon(pts2, outline=(14, 116, 144, 255), width=8)
                    draw.line([(70, 88), (128, 176), (186, 88)],
                              fill=(255, 255, 255, 255), width=28,
                              joint='curve')
                    draw.ellipse([119, 167, 137, 185],
                                 fill=(14, 116, 144, 255))
                    img.save(ico_path, sizes=[(16,16),(32,32),(48,48),
                                               (64,64),(128,128),(256,256)])
                dash_root.iconbitmap(ico_path)
                for size in (32, 64, 256):
                    im = Image.open(ico_path).resize((size, size),
                                                       Image.LANCZOS)
                    photo = ImageTk.PhotoImage(im)
                    dash_root.wm_iconphoto(True, photo)
            except Exception:
                pass

            dash_root.mainloop()
        except Exception:
            traceback.print_exc()

    # ==== Arranque con splash ====
    splash_root = tk.Tk()
    splash = SplashScreen(splash_root, lambda: launch_dashboard(splash_root))
    splash_root.mainloop()