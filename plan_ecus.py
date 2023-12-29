from pulp import *
import warnings
import pylightxl as xl
import unidecode
import sys

# Deshabilitar 'warning'
warnings.filterwarnings("ignore", message="Overwriting previously set objective.")

# Función para definir el nombre de las variables de decisión en el modelo
def nombre_var(nombre, *indices):
    return nombre+'_'+'_'.join(str(x) for x in indices)

def optVar(T, modelo, var2opt):
    n_periodos = len(T)
    modelo += lpSum((n_periodos-t+1)*var2opt[t] for t in T)
    modelo.solve()
    optVals = {t:var2opt[t].value() for t in T}
    for t in T:
        modelo += var2opt[t] <= 1.0000001*optVals[t]
    return modelo.status, value(modelo.objective)

def crea_plan(archivoDatos, archivoSolucion):
    db = xl.readxl(fn=archivoDatos, ws=('bloques','plan', 'inventarios', 'capacidad', 'constantes', 'prioridades'))
    
    # =======================================
    # Conjuntos: desde archivo 'Datos.xlsx'
    # =======================================
    
    # Bloques de días: conjunto de días dentro de los cuales algunas variables de decisión tienen el mismo valor
    B = db.ws(ws='bloques').row(row=1)[1:]
    nb = len(B)
    
    # Días: horizonte de planeación
    n_dias = db.ws(ws='bloques').row(row=3)[-1] # columnas con datos en la hoja 'plan'
    T = [dia for dia in range(1, n_dias+1)]
    
    # Bloques de días para el cloruro: B_u --> cada cambio de valor de capacidad máxima genera un bloque
    Max_u = db.ws(ws='plan').row(row=60)[1:]
    T_u = {1:[1]}
    Pmax_u = {1:Max_u[0]}
    b, j = 1, 1
    while j < len(Max_u):
        if Max_u[j] != Max_u[j-1]:
            b+=1
            T_u[b] = [j+1]
            Pmax_u[b] = Max_u[j]
        else:
            T_u[b].append(j+1)
        j+=1
    B_u = list(T_u.keys())
    
    # =======================================
    # Parámetros
    # =======================================
    
    # =======================================
    # Bloques de días: día en el que inicia cada bloque
    # =======================================
    dia_inicio = db.ws(ws='bloques').row(row=2)[1:]
    inicio = dict(zip(B,dia_inicio))
    T_j = {j:[i for i in range(inicio[j],inicio[j+1])] for j in B[:-1]}
    T_j[B[-1]] = [i for i in range(inicio[B[-1]], T[-1]+1)]
    
    # =======================================
    # Ácido clorhídrico
    # =======================================
    
    d_a = dict(zip(T, db.ws(ws='plan').row(row=4)[1:n_dias+1]))
    Io_a = db.ws(ws='inventarios').range(address='B2')[0][0]
    Ix_a = db.ws(ws='inventarios').range(address='B3')[0][0]
    Imin_a = db.ws(ws='inventarios').range(address='B4')[0][0]
    Imax_a = db.ws(ws='inventarios').range(address='B5')[0][0]
    
    Pmin_a = dict(zip(B, db.ws(ws='capacidad').row(row=3)[1:nb+1]))
    Pmax_a = dict(zip(B, db.ws(ws='capacidad').row(row=3)[6:6+nb]))
    
    ci_hcl = db.ws(ws='constantes').range(address='B4')[0][0]
    
    # =======================================
    # Hipoclorito
    # =======================================
    
    p_h15  = dict(zip(T, db.ws(ws='plan').row(row=18)[1:n_dias+1]))
    e_h25  = dict(zip(T, db.ws(ws='plan').row(row=19)[1:n_dias+1]))
    e_h525 = dict(zip(T, db.ws(ws='plan').row(row=20)[1:n_dias+1]))
    e_h83  = dict(zip(T, db.ws(ws='plan').row(row=21)[1:n_dias+1]))
    d_h15  = dict(zip(T, db.ws(ws='plan').row(row=14)[1:n_dias+1]))
    
    Io_h525 = db.ws(ws='inventarios').range(address='C2')[0][0]
    Ix_h525 = db.ws(ws='inventarios').range(address='C3')[0][0]
    Imin_h525 = db.ws(ws='inventarios').range(address='C4')[0][0]
    Imax_h525 = db.ws(ws='inventarios').range(address='C5')[0][0]
    
    Io_h15 = db.ws(ws='inventarios').range(address='D2')[0][0]
    Ix_h15 = db.ws(ws='inventarios').range(address='D3')[0][0]
    Imin_h15 = db.ws(ws='inventarios').range(address='D4')[0][0]
    Imax_h15 = db.ws(ws='inventarios').range(address='D5')[0][0]
    
    Pmin_h = dict(zip(B, db.ws(ws='capacidad').row(row=4)[1:nb+1]))
    Pmax_h = dict(zip(B, db.ws(ws='capacidad').row(row=4)[6:6+nb]))
    
    # =======================================
    # Cloro
    # =======================================
    
    s_c = db.ws(ws='constantes').range(address='B1')[0][0]
    r_c = dict(zip(T, db.ws(ws='plan').row(row=33)[1:n_dias+1]))
    
    Io_cg = db.ws(ws='inventarios').range(address='E2')[0][0]
    Ix_cg = db.ws(ws='inventarios').range(address='E3')[0][0]
    Imin_cg = db.ws(ws='inventarios').range(address='E4')[0][0]
    Imax_cg = db.ws(ws='inventarios').range(address='E5')[0][0]
    
    Io_cnt = db.ws(ws='inventarios').range(address='H2')[0][0]
    
    Pmin_c = dict(zip(B, db.ws(ws='capacidad').row(row=5)[1:nb+1]))
    Pmax_c = dict(zip(B, db.ws(ws='capacidad').row(row=5)[6:6+nb]))
    Pmax_L = dict(zip(B, db.ws(ws='capacidad').row(row=6)[6:6+nb]))
    BigM = 10000*sum([Pmax_L[j] for j in B])
    
    # =======================================
    # Soda
    # =======================================
    
    d_s  = dict(zip(T, db.ws(ws='plan').row(row=43)[1:n_dias+1]))
    Io_s = db.ws(ws='inventarios').range(address='F2')[0][0]
    Ix_s = db.ws(ws='inventarios').range(address='F3')[0][0]
    Imin_s = db.ws(ws='inventarios').range(address='F4')[0][0]
    Imax_s = db.ws(ws='inventarios').range(address='F5')[0][0]
    ci_NaOH = db.ws(ws='constantes').range(address='B5')[0][0]
    
    # =======================================
    # Cloruro
    # =======================================
    
    d_u  = dict(zip(T, db.ws(ws='plan').row(row=53)[1:n_dias+1]))
    e_u  = dict(zip(T, db.ws(ws='plan').row(row=56)[1:n_dias+1]))
    
    Io_u = db.ws(ws='inventarios').range(address='G2')[0][0]
    Ix_u = db.ws(ws='inventarios').range(address='G3')[0][0]
    Imin_u = db.ws(ws='inventarios').range(address='G4')[0][0]
    Imax_u = db.ws(ws='inventarios').range(address='G5')[0][0]
    
    # =======================================
    # ECUS
    # =======================================
    Qmax = dict(zip(B, db.ws(ws='bloques').row(row=4)[1:]))
    
    # =======================================
    # Otras constantes
    # =======================================
    b = db.ws(ws='constantes').range(address='B2')[0][0]
    g = db.ws(ws='constantes').range(address='B3')[0][0]
    
    # =======================================
    # Variables de decisión
    # =======================================
    
    # ====================================================================================
    # Cantidades a producir
    # ====================================================================================
    x_hI = {t:LpVariable(nombre_var('x_hI',t), 0, None, LpContinuous) for t in T}
    x_hL = {t:LpVariable(nombre_var('x_hL',t), 0, None, LpContinuous) for t in T}
    x_a  = {t:LpVariable(nombre_var('x_a',t),  0, None, LpContinuous) for t in T}
    x_u  = {t:LpVariable(nombre_var('x_u',t),  0, None, LpContinuous) for t in T}
    x_s  = {t:LpVariable(nombre_var('x_s',t),  0, None, LpContinuous) for t in T}
    x_c  = {t:LpVariable(nombre_var('x_c',t),  0, None, LpContinuous) for t in T}
    L_c  = {t:LpVariable(nombre_var('L_c',t),  0, None, LpInteger) for t in T}  # Número de contenedores a llenar.
    w    = {t:LpVariable(nombre_var('w',t), cat=LpBinary) for t in T}  # Máximo número de contenenedores a llenar --> min{I_v, Pmax_L}
    
    # ====================================================================================
    # Consumos internos   y = consumos internos posibles
    # ====================================================================================
    y_h25  = {t:LpVariable(nombre_var('y_h25',t),  0, None, LpContinuous) for t in T}
    y_h525 = {t:LpVariable(nombre_var('y_h525',t), 0, None, LpContinuous) for t in T}
    y_h83  = {t:LpVariable(nombre_var('y_h83',t),  0, None, LpContinuous) for t in T}
    y_s    = {t:LpVariable(nombre_var('y_s',t),    0, None, LpContinuous) for t in T}
    y_a    = {t:LpVariable(nombre_var('y_a',t),    0, None, LpContinuous) for t in T}
    y_u    = {t:LpVariable(nombre_var('y_u',t),    0, None, LpContinuous) for t in T}
    # ====================================================================================
    v_y_h25  = {t:LpVariable(nombre_var('v_y_h25',t),  0, None, LpContinuous) for t in T}  # violación de la meta de consumo interno de hipoclorito al 2.5%
    v_y_h525 = {t:LpVariable(nombre_var('v_y_h525',t), 0, None, LpContinuous) for t in T}  # violación de la meta de consumo interno de hipoclorito al 5.25%
    v_y_h83  = {t:LpVariable(nombre_var('v_y_h83',t),  0, None, LpContinuous) for t in T}  # violación de la meta de consumo interno de hipoclorito al 8.3%
    v_y_u    = {t:LpVariable(nombre_var('v_y_u',t),    0, None, LpContinuous) for t in T}  # violación de la meta de consumo interno de cloruro
    # ====================================================================================
    
    # ====================================================================================
    # Ventas  z = posibles
    # ====================================================================================
    z_h15 = {t:LpVariable(nombre_var('z_h15',t), 0, None, LpContinuous) for t in T}
    z_a   = {t:LpVariable(nombre_var('z_a',t),   0, None, LpContinuous) for t in T}
    z_u   = {t:LpVariable(nombre_var('z_u',t),   0, None, LpContinuous) for t in T}
    z_s   = {t:LpVariable(nombre_var('z_s',t),   0, None, LpContinuous) for t in T}
    # ====================================================================================
    v_z_h15 = {t:LpVariable(nombre_var('v_z_h15',t), 0, None, LpContinuous) for t in T}  # violación de la meta de venta de hipoclorito al 15%
    v_z_a   = {t:LpVariable(nombre_var('v_z_a',t),   0, None, LpContinuous) for t in T}  # violación de la meta de venta de ácido clorhídrico
    v_z_u   = {t:LpVariable(nombre_var('v_z_u',t),   0, None, LpContinuous) for t in T}  # violación de la meta de venta de ácido cloruro
    v_z_s   = {t:LpVariable(nombre_var('v_z_s',t),   0, None, LpContinuous) for t in T}  # violación de la meta de venta de soda
    v_L_c   = {t:LpVariable(nombre_var('v_L_c',t),   0, None, LpContinuous) for t in T}  # violación de la meta de envasado de cloro en contenedores
    # ====================================================================================
    
    # ====================================================================================
    # Inventarios al final del dia
    # ====================================================================================
    I_h15  = {t:LpVariable(nombre_var('I_h15',t),  0, None, LpContinuous) for t in T}
    I_h525 = {t:LpVariable(nombre_var('I_h525',t), 0, None, LpContinuous) for t in T}
    I_a    = {t:LpVariable(nombre_var('I_a',t),    0, None, LpContinuous) for t in T}
    I_u    = {t:LpVariable(nombre_var('I_u',t),    0, None, LpContinuous) for t in T}
    I_s    = {t:LpVariable(nombre_var('I_s',t),    0, None, LpContinuous) for t in T}
    I_cg   = {t:LpVariable(nombre_var('I_cg',t),   0, None, LpContinuous) for t in T}
    I_v    = {t:LpVariable(nombre_var('I_v',t),    0, None, LpContinuous) for t in T}   # inventario de contenedores al inicio del dia
    # ====================================================================================
    v_I_h15  = {t:LpVariable(nombre_var('v_I_h15',t),  0, None, LpContinuous) for t in T}   # violación de la meta de inventario mínimo de hipoclorito al 15%
    v_I_h525 = {t:LpVariable(nombre_var('v_I_h525',t), 0, None, LpContinuous) for t in T}   # violación de la meta de inventario mínimo de hipoclorito al 5.25%
    v_I_a    = {t:LpVariable(nombre_var('v_I_a',t),    0, None, LpContinuous) for t in T}      # violación de la meta de inventario mínimo de ácido clorhídrico
    v_I_u    = {t:LpVariable(nombre_var('v_I_u',t),    0, None, LpContinuous) for t in T}      # violación de la meta de inventario mínimo de cloruro
    v_I_s    = {t:LpVariable(nombre_var('v_I_s',t),    0, None, LpContinuous) for t in T}      # violación de la meta de inventario mínimo de soda
    v_I_cg   = {t:LpVariable(nombre_var('v_I_cg',t),   0, None, LpContinuous) for t in T}      # violación de la meta de inventario mínimo de clor a granel
    # ====================================================================================
    
    # ====================================================================================
    # Otras
    # ====================================================================================
    q = {t:LpVariable(nombre_var('q',t), 0, None, LpContinuous) for t in T}  # q = ecus
    v_q = {t:LpVariable(nombre_var('v_q',t), 0, None, LpContinuous) for t in T}  # violación de la meta de ecus
    m = {t:LpVariable(nombre_var('m',t), 0, None, LpInteger) for t in T}     # m Multiplo de 30 para producción de hipo
    u24 = {j:LpVariable(nombre_var('u24',j), cat=LpBinary) for j in B_u}   # Variable binaria que activa producción de cloruro 24 toneladas
    u42 = {j:LpVariable(nombre_var('u42',j), cat=LpBinary) for j in B_u}   # Variable binaria que activa producción de cloruro 42 toneladas
    
    # Modelo
    modelo_ecus = LpProblem('MaxECUS', LpMinimize)
    
    # Restricciones
    
    # Balance de masa hipoclorito al 15%
    modelo_ecus += I_h15[1] == Io_h15 + x_hI[1] - z_h15[1] - y_h25[1] - y_h83[1]
    for t in T[1:]:
        modelo_ecus += I_h15[t] == I_h15[t-1] + x_hI[t] - z_h15[t] - y_h25[t] - y_h83[t]
    
    # Balance de masa hipoclorito al 5.25%
    modelo_ecus += I_h525[1] == Io_h525 + x_hL[1] + p_h15[1] - y_h525[1]
    for t in T[1:]:
        modelo_ecus += I_h525[t] == I_h525[t-1] + x_hL[t] + p_h15[t] - y_h525[t]
    
    # Balance de masa ácido clorhídrico
    modelo_ecus += I_a[1] == Io_a + x_a[1] - y_a[1] - z_a[1]
    for t in T[1:]:
        modelo_ecus += I_a[t] == I_a[t-1] + x_a[t] - y_a[t] - z_a[t]
    
    # Balance de masa de soda
    modelo_ecus += I_s[1] == Io_s + x_s[1] - y_s[1] - z_s[1]
    for t in T[1:]:
        modelo_ecus += I_s[t] == I_s[t-1] + x_s[t] - y_s[t] - z_s[t]
    
    # Balance de masa cloruro
    modelo_ecus += I_u[1] == Io_u + x_u[1] - y_u[1] - z_u[1]
    for t in T[1:]:
        modelo_ecus += I_u[t] == I_u[t-1] + x_u[t] - y_u[t] - z_u[t]
    
    # Balance de masa de cloro a granel
    modelo_ecus += I_cg[1] == Io_cg + x_c[1] - s_c*L_c[1]
    for t in T[1:]:
        modelo_ecus += I_cg[t] == I_cg[t-1] + x_c[t] - s_c*L_c[t]
    
    # Balance de masa de contenedores de cloro vacíos
    modelo_ecus += I_v[1] == Io_cnt + r_c[1]
    for t in T[1:]:
        modelo_ecus += I_v[t] == I_v[t-1] - L_c[t-1] + r_c[t]
    
    # Restricción de proceso: relación entre ecus y producción de soda
    for t in T:
        modelo_ecus += x_s[t] == 2.256*q[t]
    
    # Restricción de proceso: producción de cloruro
    for j in B_u:
        for t in T_u[j]:
            modelo_ecus += x_u[t] == 24*u24[j] + 42*u42[j]
    
    for j in B_u:
        modelo_ecus += u24[j] + u42[j] <= 1
    
    # Restricción de proceso: producción de ecus
    for t in T:
        modelo_ecus += q[t] == 0.325*x_a[t] + 0.144*(x_hI[t] + x_hL[t] + p_h15[t]) + x_c[t]
    
    # ==============================================================================
    # Restricción de balance: Producción de hipoclorito industrial igual por bloque
    for j in B:
        for t in T_j[j][1:]:
            modelo_ecus += x_hI[t] == x_hI[t-1]
    
    # Restricción de balance: Producción de hipoclorito en lotes igual por bloque
    for j in B:
        for t in T_j[j][1:]:
            modelo_ecus += x_hL[t] == x_hL[t-1]
    
    # Restricción de balance: Producción de ácido igual por bloque
    for j in B:
        for t in T_j[j][1:]:
            modelo_ecus += x_a[t] == x_a[t-1]
    
    # Restricción de balance: Producción de cloro igual por bloque
    for j in B:
        for t in T_j[j][1:]:
            modelo_ecus += x_c[t] == x_c[t-1]
    
    # Restricción de balance: Producción de cloruro igual por bloque
    for j in B_u:
        for t in T_u[j][1:]:
            modelo_ecus += x_u[t] == x_u[t-1]
    # ==============================================================================
    
    # Restricción de proceso: consumo interno de ácido clorhídrico
    for t in T:
        modelo_ecus += y_a[t] == ci_hcl + b*x_u[t]
    
    # Restricción de proceso: consumo interno de soda
    for t in T:
        modelo_ecus += y_s[t] == ci_NaOH + g*(x_hI[t] + x_hL[t] + p_h15[t])
    
    # Restricción de proceso: producción de hipoclorito en múltiplos de 30
    for t in T:
        modelo_ecus += 30*m[t] == x_hI[t] + x_hL[t]
    
    # Restrición de proceso: hipoclorito en lotes
    modelo_ecus += y_h525[1] + y_h25[1] + y_h83[1] <= x_hL[1] + p_h15[1] + Io_h525
    for t in T[1:]:
        modelo_ecus += y_h525[t] + y_h25[t] + y_h83[t] <= x_hL[t] + p_h15[t] + I_h525[t-1]
    
    # Límite al consumo interno de hipoclorito al 2.5%
    for t in T:
        modelo_ecus += y_h25[t] + v_y_h25[t] == e_h25[t]
    
    # Límite al consumo interno de hipoclorito al 5.25%
    for t in T:
        modelo_ecus += y_h525[t] + v_y_h525[t] == e_h525[t]
    
    # Límite al consumo interno de hipoclorito al 8.3%
    for t in T:
        modelo_ecus += y_h83[t] + v_y_h83[t] == e_h83[t]
    
    # Límite al consumo interno de cloruro
    for t in T:
        modelo_ecus += y_u[t] + v_y_u[t] == e_u[t]
    
    # Límite a las ventas de ácido clorhídrico
    for t in T:
        modelo_ecus += z_a[t] + v_z_a[t] == d_a[t]
    
    # Límite a las ventas de cloruro
    for t in T:
        modelo_ecus += z_u[t] + v_z_u[t] == d_u[t]
    
    # Límite a las ventas de hipoclorito al 15%
    for t in T:
        modelo_ecus += z_h15[t] + v_z_h15[t] == d_h15[t]
    
    # Límite a las ventas de soda
    for t in T:
        modelo_ecus += z_s[t] + v_z_s[t] == d_s[t]
    
    # Límite a la cantidad de contenedores de cloro a envasar
    for j in B:
      for t in T_j[j]:
        modelo_ecus += I_v[t] >= w[t]*Pmax_L[j]
        modelo_ecus += Pmax_L[j] >= I_v[t] -  BigM*w[t]
        modelo_ecus += v_L_c[t] >= Pmax_L[j] - L_c[t] - BigM*(1 - w[t])
        modelo_ecus += v_L_c[t] >= I_v[t] - L_c[t] - BigM*w[t]

    # Límite a la máxima producción de ecus
    for j in B:
        for t in T_j[j]:
            modelo_ecus += q[t] + v_q[t] == Qmax[j]
    
    # Límites a la producción de ácido clorhídrio
    for j in B:
      for t in T_j[j]:
          modelo_ecus += x_a[t] >= Pmin_a[j]
          modelo_ecus += x_a[t] <= Pmax_a[j]
    
    # Límites a la producción de hipoclorito
    for j in B:
      for t in T_j[j]:
        modelo_ecus += x_hI[t] + x_hL[t] >= Pmin_h[j]
        modelo_ecus += x_hI[t] + x_hL[t] <= Pmax_h[j]
    
    # Límites a la producción de cloro
    for j in B:
      for t in T_j[j]:
        modelo_ecus += x_c[t] >= Pmin_c[j]
        modelo_ecus += x_c[t] <= Pmax_c[j]
        modelo_ecus += L_c[t] <= Pmax_L[j]
        modelo_ecus += L_c[t] <= I_v[t]
    
    # Límites a la producción de cloruro
    for j in B_u:
      for t in T_u[j]:
          modelo_ecus += x_u[t] <= Pmax_u[j]

    # Inventarios mínimos
    for t in T:
        modelo_ecus += I_h525[t] + v_I_h525[t] >= Imin_h525
        modelo_ecus += I_h15[t]  + v_I_h15[t]  >= Imin_h15
        modelo_ecus += I_cg[t]   + v_I_cg[t]   >= Imin_cg
        modelo_ecus += I_a[t]    + v_I_a[t]    >= Imin_a
        modelo_ecus += I_s[t]    + v_I_s[t]    >= Imin_s
        modelo_ecus += I_u[t]    + v_I_u[t]    >= Imin_u
    
    # Inventarios críticos
    for t in T:
        modelo_ecus += v_I_h525[t] <= Imin_h525 - Ix_h525
        modelo_ecus += v_I_h15[t]  <= Imin_h15 - Ix_h15
        modelo_ecus += v_I_cg[t]   <= Imin_cg - Ix_cg
        modelo_ecus += v_I_a[t]    <= Imin_a - Ix_a
        modelo_ecus += v_I_s[t]    <= Imin_s - Ix_s
        modelo_ecus += v_I_u[t]    <= Imin_u - Ix_u
    
    # Capacidades de almacenamiento
    for t in T:
        modelo_ecus += I_a[t]    <= Imax_a
        modelo_ecus += I_h525[t] <= Imax_h525
        modelo_ecus += I_h15[t]  <= Imax_h15
        modelo_ecus += I_cg[t]   <= Imax_cg
        modelo_ecus += I_s[t]    <= Imax_s
        modelo_ecus += I_u[t]    <= Imax_u

    # ================
    # OBJETIVOS
    # ================
    objetivos = [j[0] for j in db.ws(ws='prioridades').range(address='A2:A17')]
    prioridades = [j[0]-1 for j in db.ws(ws='prioridades').range(address='B2:B17')]
    
    orden = {i:None for i in prioridades}
    
    for i in prioridades:
      nombre_objetivo = unidecode.unidecode(objetivos[i]).lower()
      if 'ecus' in nombre_objetivo:
        orden[i] = v_q
      if 'venta' in nombre_objetivo:
        if '15%' in nombre_objetivo:
          orden[i] = v_z_h15
        elif 'acido' in nombre_objetivo:
          orden[i] = v_z_a
        elif 'soda' in nombre_objetivo:
          orden[i] = v_z_s
        elif 'cloruro' in nombre_objetivo:
          orden[i] = v_z_u
      if 'envasado' in nombre_objetivo and 'contenedores' in nombre_objetivo:
        orden[i] = v_L_c
      if 'consumo' in nombre_objetivo and 'interno' in nombre_objetivo:
        if '2.5%' in nombre_objetivo:
          orden[i] = v_y_h25
        elif '8.3%' in nombre_objetivo:
          orden[i] = v_y_h83
        elif '5.25%' in nombre_objetivo:
          orden[i] = v_y_h525
        elif 'cloruro' in nombre_objetivo:
          orden[i] = v_y_u
      if 'inventario' in nombre_objetivo:
        if '15%' in nombre_objetivo:
          orden[i] = v_I_h15
        elif '5.25%' in nombre_objetivo:
          orden[i] = v_I_h525
        elif 'cloro' in nombre_objetivo:
          orden[i] = v_I_cg
        elif 'soda' in nombre_objetivo:
          orden[i] = v_I_s
        elif 'acido' in nombre_objetivo:
          orden[i] = v_I_a
        elif 'cloruro' in nombre_objetivo:
          orden[i] = v_I_u

    # ===========================================================================================
    # Resolver en 16 fases: el orden en que se llama la función optVar determina las prioridades
    # ===========================================================================================
    for i in orden:
      resultado, obj_opt = optVar(T, modelo_ecus, orden[i])
      if resultado == 1:
          print(f'Solución óptima: El incumplimiento total en la variable {orden[i]} fue de {round(obj_opt,2)}')
          #db.ws(ws='plan').update_index(row=1, col=1, val='Solución Óptima')
      else:
          print(f'la optimziación de la variable {orden[i]} resultó infactible')
          db.ws(ws='plan').update_index(row=1, col=1, val='Problema Infactible '+str(i))
          break

    # =======================================
    # Escribe solución en archivo 'xlsx'
    # =======================================
    invIni_h15  = [Io_h15]+[I_h15[t].value() for t in T][:-1]
    invIni_h525 = [Io_h525]+[I_h525[t].value() for t in T][:-1]
    invIni_a = [Io_a]+[I_a[t].value() for t in T][:-1]
    invIni_u = [Io_u]+[I_u[t].value() for t in T][:-1]
    invIni_s = [Io_s]+[I_s[t].value() for t in T][:-1]
    invIni_cg = [Io_cg]+[I_cg[t].value() for t in T][:-1]
    
    for i in range(len(T)):
      # Cantidades a producir
      db.ws(ws='plan').update_index(row=16, col=i+2, val=[round(x_hI[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=17, col=i+2, val=[round(x_hL[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=6,  col=i+2, val=[round(x_a[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=55, col=i+2, val=[round(x_u[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=45, col=i+2, val=[round(x_s[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=34, col=i+2, val=[round(x_c[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=36, col=i+2, val=[L_c[t].value() for t in T][i])
      # Consumos internos
      db.ws(ws='plan').update_index(row=22, col=i+2, val=[round(y_h25[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=23, col=i+2, val=[round(y_h525[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=24, col=i+2, val=[round(y_h83[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=57, col=i+2, val=[round(y_u[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=46, col=i+2, val=[round(y_s[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=7,  col=i+2, val=[round(y_a[t].value(),2) for t in T][i])
      # Ventas
      db.ws(ws='plan').update_index(row=15, col=i+2, val=[round(z_h15[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=5,  col=i+2, val=[round(z_a[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=54, col=i+2, val=[round(z_u[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=44, col=i+2, val=[round(z_s[t].value(),2) for t in T][i])
      # Inventarios iniciales
      db.ws(ws='plan').update_index(row=25, col=i+2, val=round(invIni_h15[i],2))
      db.ws(ws='plan').update_index(row=27, col=i+2, val=round(invIni_h525[i],2))
      db.ws(ws='plan').update_index(row=8,  col=i+2, val=round(invIni_a[i],2))
      db.ws(ws='plan').update_index(row=58, col=i+2, val=round(invIni_u[i],2))
      db.ws(ws='plan').update_index(row=47, col=i+2, val=round(invIni_s[i],2))
      db.ws(ws='plan').update_index(row=37, col=i+2, val=round(invIni_cg[i],2))
      # Inventarios finales
      db.ws(ws='plan').update_index(row=26, col=i+2, val=[round(I_h15[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=28, col=i+2, val=[round(I_h525[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=9,  col=i+2, val=[round(I_a[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=59, col=i+2, val=[round(I_u[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=48, col=i+2, val=[round(I_s[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=35, col=i+2, val=[round(I_v[t].value(),2) for t in T][i])
      db.ws(ws='plan').update_index(row=38, col=i+2, val=[round(I_cg[t].value(),2) for t in T][i])
      # Ecus
      db.ws(ws='plan').update_index(row=64, col=i+2, val=[round(q[t].value(),2) for t in T][i])
    xl.writexl(db=db, fn=archivoSolucion)
    
if __name__== "__main__":
    crea_plan(sys.argv[1], sys.argv[2])