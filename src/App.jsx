import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { kvGet, kvSet, supabase } from './supabaseClient';
import { 
  BookOpen, Book, FileText, BarChart2, PieChart, 
  Sun, Moon, Plus, Trash2, Save, LayoutDashboard, 
  Settings, CheckCircle, AlertCircle, Package, 
  DownloadCloud, Loader2, Database, FileSpreadsheet,
  RefreshCw, AlertTriangle, X, Edit3
} from 'lucide-react';

const DEFAULT_ACCOUNTS = [
  // 1. ACTIVO
  { id: 'acc_1', code: '1', name: 'ACTIVO', type: 'Activo', category: 'Titulo', nature: 'Deudora' },
  { id: 'acc_2', code: '1.1.0.0', name: 'ACTIVO CIRCULANTE', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_3', code: '1.1.01', name: 'EFECTIVO Y SUS EQUIVALENTES', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_4', code: '1.1.01.001', name: 'Efectivo en caja', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_5', code: '1.1.02', name: 'BANCOS', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_6', code: '1.1.02.001', name: 'BANCO BANESCO', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_7', code: '1.1.03', name: 'CUENTAS POR COBRAR', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_8', code: '1.1.03.001', name: 'Cuentas por cobrar clientes', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_9', code: '1.1.03.002', name: 'Cuentas por cobrar accionistas', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_10', code: '1.1.03.003', name: 'Cuentas por cobrar MPPFNS', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_11', code: '1.1.03.004', name: 'Cuentas por cobrar MPPT', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_12', code: '1.1.03.005', name: 'Cuentas por cobrar empleados', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_13', code: '1.1.03.006', name: 'Cuentas por cobrar intercompañias', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_14', code: '1.1.04', name: 'INVENTARIO', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_15', code: '1.1.04.001', name: 'Inventario de mercancía', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_16', code: '1.1.05', name: 'OTRAS CUENTAS POR COBRAR', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_17', code: '1.1.05.001', name: 'Anticipo a proveedores', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_18', code: '1.1.06', name: 'PREPAGADOS', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_19', code: '1.1.06.001', name: 'Pólizas de seguro', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_20', code: '1.1.07', name: 'CRÉDITOS FISCALES', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_21', code: '1.1.07.001', name: 'Impuesto Sobre La Renta', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_22', code: '1.1.07.002', name: 'Impuesto pagado por anticipado', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_23', code: '1.2.0.0', name: 'ACTIVO NO CIRCULANTE', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_24', code: '1.2.01', name: 'ACTIVO FIJO', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_25', code: '1.2.01.001', name: 'HERRAMIENTAS', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_26', code: '1.2.01.002', name: 'EQUIPO DE COMPUTACIÓN', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_27', code: '1.2.02', name: 'DEPRECIACIÓN ACUMULADA ACTIVO FIJO', type: 'Activo', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_28', code: '1.2.02.001', name: 'DEPREC ACUMULADA HERRAMIENTAS', type: 'Activo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_29', code: '1.2.02.002', name: 'DEPREC ACUMULADA EQUIP COMPUTACIÓN', type: 'Activo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_30', code: '1.2.03', name: 'OTROS ACTIVOS', type: 'Activo', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_31', code: '1.2.03.001', name: 'Iva Crédito Fiscal (COMPRAS)', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_32', code: '1.2.03.002', name: 'Excedente de Crédito Fiscal', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_33', code: '1.2.03.003', name: 'Retención Actividades Económicas', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_34', code: '1.2.03.004', name: 'Iva retenido por clientes', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_35', code: '1.2.03.005', name: 'Retención de ISLR', type: 'Activo', category: 'Cuenta', nature: 'Deudora' },

  // 2. PASIVO
  { id: 'acc_36', code: '2', name: 'PASIVO', type: 'Pasivo', category: 'Titulo', nature: 'Acreedora' },
  { id: 'acc_37', code: '2.1.0.0', name: 'PASIVO A CORTO PLAZO', type: 'Pasivo', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_38', code: '2.1.01', name: 'CUENTAS POR PAGAR A CORTO PLAZO', type: 'Pasivo', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_39', code: '2.1.01.001', name: 'CUENTAS POR PAGAR ACCIONISTA', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_40', code: '2.1.01.002', name: 'CUENTAS POR PAGAR A PROVEEDORES', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_41', code: '2.1.01.003', name: 'CUENTAS POR PAGAR INTERCOMPAÑIAS', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_42', code: '2.1.01.004', name: 'CUENTAS POR PAGAR EMPLEADOS', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_43', code: '2.1.02', name: 'RETENCIONES E IMPUESTOS', type: 'Pasivo', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_44', code: '2.1.02.001', name: 'IMPUESTO SOBRE LA RENTA', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_45', code: '2.1.02.002', name: 'Iva Débito Fiscal (VENTAS)', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_46', code: '2.1.02.003', name: 'RET PATENTE DE INDUSTRIA Y COMERCIO', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_47', code: '2.1.02.004', name: 'IVA RETENIDO A PROVEEDORES', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_48', code: '2.1.02.005', name: 'IVA POR PAGAR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_49', code: '2.1.03', name: 'BENEFICIOS AL PERSONAL POR PAGAR', type: 'Pasivo', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_50', code: '2.1.03.001', name: 'NÓMINA POR PAGAR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_51', code: '2.1.03.002', name: 'IVSS POR PAGAR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_52', code: '2.1.03.003', name: 'RPE POR PAGAR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_53', code: '2.1.03.004', name: 'BANAVIH POR PAGAR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_54', code: '2.1.03.005', name: 'INCES POR PAGAR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_55', code: '2.1.03.006', name: 'PROVISIÓN DE VACACIONES', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_56', code: '2.1.03.007', name: 'PROVISIÓN DE BONO VACACIONAL', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_57', code: '2.1.03.008', name: 'PROVISIÓN DE UTILIDADES', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_58', code: '2.1.03.009', name: 'PROVISIÓN DE PRESTACIONES SOCIALES', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_59', code: '2.1.03.010', name: 'PROVISIÓN DE INTERESES SOBRE PRESTACIONES SOCIALES', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_60', code: '2.1.04', name: 'RETENCIONES POR PAGAR', type: 'Pasivo', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_61', code: '2.1.04.001', name: 'BANAVIH RETENCIÓN TRABAJADOR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_62', code: '2.1.04.002', name: 'IVSS RETENCIÓN TRABAJADOR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_63', code: '2.1.04.003', name: 'RPE RETENCIÓN TRABAJADOR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_64', code: '2.1.04.004', name: 'INCES RETENCIÓN TRABAJADOR', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_65', code: '2.1.05', name: 'OTRAS CUENTAS POR PAGAR', type: 'Pasivo', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_66', code: '2.1.05.001', name: 'ANTICIPO RECIBIDO DE CLIENTE MPPT', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_67', code: '2.1.05.002', name: 'ANTICIPO RECIBIDO DE CLIENTE FNS', type: 'Pasivo', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_68', code: '2.2.0.0', name: 'PASIVO A LARGO PLAZO', type: 'Pasivo', category: 'Grupo', nature: 'Acreedora' },

  // 3. PATRIMONIO
  { id: 'acc_69', code: '3', name: 'PATRIMONIO', type: 'Patrimonio', category: 'Titulo', nature: 'Acreedora' },
  { id: 'acc_70', code: '3.1.0.0', name: 'CAPITAL', type: 'Patrimonio', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_71', code: '3.1.01', name: 'CAPITAL SOCIAL', type: 'Patrimonio', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_72', code: '3.1.01.001', name: 'CAPITAL SOCIAL', type: 'Patrimonio', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_73', code: '3.1.01.002', name: 'APORTE DE ACCIONISTAS', type: 'Patrimonio', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_74', code: '3.2.0.0', name: 'RESERVAS', type: 'Patrimonio', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_75', code: '3.2.01', name: 'RESERVAS DE CAPITAL', type: 'Patrimonio', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_76', code: '3.2.01.001', name: 'RESERVA LEGAL', type: 'Patrimonio', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_77', code: '3.3.0.0', name: 'SUPERAVIT', type: 'Patrimonio', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_78', code: '3.3.01', name: 'SUPERAVIT ACUMULADO', type: 'Patrimonio', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_79', code: '3.3.01.001', name: 'UTILIDADES NO DISTRIBUIDAS', type: 'Patrimonio', category: 'Cuenta', nature: 'Acreedora' },

  // 4. INGRESOS
  { id: 'acc_80', code: '4', name: 'INGRESOS', type: 'Ingreso', category: 'Titulo', nature: 'Acreedora' },
  { id: 'acc_81', code: '4.1.0.0', name: 'INGRESOS OPERATIVOS', type: 'Ingreso', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_82', code: '4.1.01', name: 'INGRESOS POR SERVICIOS', type: 'Ingreso', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_83', code: '4.1.01.001', name: 'INGRESOS POR SERVICIO DE MANTENIMIENTO DE ASCENSORES', type: 'Ingreso', category: 'Cuenta', nature: 'Acreedora' },

  // 5. COSTOS
  { id: 'acc_84', code: '5', name: 'COSTOS', type: 'Gasto', category: 'Titulo', nature: 'Deudora' },
  { id: 'acc_85', code: '5.1.0.0', name: 'COSTOS DE VENTAS Y SERVICIOS', type: 'Gasto', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_86', code: '5.1.01', name: 'COSTO DE COMPRAS', type: 'Gasto', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_87', code: '5.1.01.001', name: 'COSTOS DE MATERIAL', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_88', code: '5.1.01.002', name: 'COSTOS DE MANO DE OBRA', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },

  // 6. GASTOS
  { id: 'acc_89', code: '6', name: 'GASTOS', type: 'Gasto', category: 'Titulo', nature: 'Deudora' },
  { id: 'acc_90', code: '6.1.0.0', name: 'GASTOS DE ADMINISTRACIÓN', type: 'Gasto', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_91', code: '6.1.01', name: 'GASTOS DE ADMINISTRACIÓN DEL PERSONAL', type: 'Gasto', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_92', code: '6.1.01.001', name: 'Sueldos y Salarios', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_93', code: '6.1.01.002', name: 'Banavih Aporte Patrono', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_94', code: '6.1.01.003', name: 'IVSS Aporte Patrono', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_95', code: '6.1.01.004', name: 'RPE Aporte Patrono', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_96', code: '6.1.01.005', name: 'Inces Aporte Patrono', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_97', code: '6.1.01.006', name: 'Gasto de prestaciones sociales', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_98', code: '6.1.01.007', name: 'Gasto de Vacaciones', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_99', code: '6.1.01.008', name: 'Gasto de Bono Vacacional', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_100', code: '6.1.01.009', name: 'Gasto de Utilidades', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_101', code: '6.1.01.010', name: 'Gasto de intereses sobre prestaciones sociales', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_102', code: '6.1.01.011', name: 'Gasto de Formación profesional', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_103', code: '6.1.02', name: 'GASTO DE DEPRECIACIÓN Y AMORTIZACIONES', type: 'Gasto', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_104', code: '6.1.02.001', name: 'GASTO DE DEPRECIACIÓN HERRAMIENTAS', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_105', code: '6.1.02.002', name: 'GASTO DE DEPRECIACIÓN EQUIPO DE COMPUTACIÓN', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_106', code: '6.1.03', name: 'GASTOS DE IMPUESTOS', type: 'Gasto', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_107', code: '6.1.03.001', name: 'PATENTE DE INDUSTRIA Y COMERCIO', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_108', code: '6.1.03.002', name: 'OTROS IMPUESTOS MUNICIPALES', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_109', code: '6.1.03.003', name: 'RETENCIÓN 1*1000', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_110', code: '6.1.03.004', name: 'RESPONSABILIDAD SOCIAL', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_111', code: '6.2.0.0', name: 'GASTOS DE OPERACIÓN', type: 'Gasto', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_112', code: '6.2.01.001', name: 'GASTOS DE MANTENIMIENTO Y REPARACIÓN', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_113', code: '6.2.01.002', name: 'GASTOS DE REGISTRO DE EMPRESA', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_114', code: '6.2.01.003', name: 'SAREN', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_115', code: '6.2.01.004', name: 'COMIDAS Y REFRIGERIOS', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_116', code: '6.2.01.005', name: 'GASTO DE TELÉFONO', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_117', code: '6.2.01.006', name: 'GASTO DE LUZ', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_118', code: '6.2.01.007', name: 'INTERNET', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_119', code: '6.2.01.008', name: 'CONDOMINIO', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_120', code: '6.2.01.009', name: 'FIANZA DE FIEL CUMPLIMIENTO', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_121', code: '6.2.01.010', name: 'FIANZA DE ANTICIPO', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_122', code: '6.2.01.011', name: 'FIANZA LABORAL', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_123', code: '6.2.01.012', name: 'GASTO DE MATERIAL DE OFICINA', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_124', code: '6.2.01.013', name: 'MOVILIZACIONES', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_125', code: '6.3.0.0', name: 'GASTOS NO OPERATIVOS', type: 'Gasto', category: 'Grupo', nature: 'Deudora' },
  { id: 'acc_126', code: '6.3.01.001', name: 'INTERESES MORATORIOS', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_127', code: '6.3.01.002', name: 'GASTOS POR MULTAS Y SANCIONES', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },

  // 7. INGRESOS Y EGRESOS FINANCIEROS
  { id: 'acc_128', code: '7', name: 'INGRESOS Y EGRESOS FINANCIEROS', type: 'Ingreso', category: 'Titulo', nature: 'Acreedora' },
  { id: 'acc_129', code: '7.1.0.0', name: 'FINANCIEROS', type: 'Ingreso', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_130', code: '7.1.01.001', name: 'INGRESOS POR INTERESES BANCARIOS', type: 'Ingreso', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_131', code: '7.1.01.002', name: 'INGRESO POR DIFERENCIA EN CAMBIO', type: 'Ingreso', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_132', code: '7.1.02.001', name: 'COMISIONES BANCARIAS', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_133', code: '7.1.02.002', name: 'GASTOS BANCARIOS', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_134', code: '7.1.02.003', name: 'EGRESOS POR DIFERENCIA EN CAMBIO', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },

  // 8. EXTRAORDINARIOS
  { id: 'acc_135', code: '8', name: 'INGRESOS Y EGRESOS EXTRAORDINARIOS', type: 'Ingreso', category: 'Titulo', nature: 'Acreedora' },
  { id: 'acc_136', code: '8.1.0.0', name: 'EXTRAORDINARIOS', type: 'Ingreso', category: 'Grupo', nature: 'Acreedora' },
  { id: 'acc_137', code: '8.1.01.001', name: 'INGRESOS VARIOS', type: 'Ingreso', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_138', code: '8.1.01.002', name: 'GANANCIA EN VENTA DE ACTIVOS FIJOS', type: 'Ingreso', category: 'Cuenta', nature: 'Acreedora' },
  { id: 'acc_139', code: '8.1.02.001', name: 'EGRESOS VARIOS', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' },
  { id: 'acc_140', code: '8.1.02.002', name: 'PERDIDA EN VENTA DE ACTIVOS FIJOS', type: 'Gasto', category: 'Cuenta', nature: 'Deudora' }
];

const DEFAULT_INVENTORY = [];

const formatCurrency = (amount, currency = 'VES', exchangeRate = 1) => {
  const converted = currency === 'USD' ? (amount / (exchangeRate || 1)) : amount;
  const currSymbol = currency === 'USD' ? 'USD' : 'VES';
  const locale = currency === 'USD' ? 'en-US' : 'es-VE';
  return new Intl.NumberFormat(locale, { style: 'currency', currency: currSymbol }).format(converted || 0);
};

const calculateBalance = (debits, credits, nature) => {
  return nature === 'Deudora' ? debits - credits : credits - debits;
};

// Construye y descarga un archivo .xlsx real (no CSV) a partir de una o varias hojas.
// Cada hoja es { name: 'Nombre pestaña', data: [[fila1...], [fila2...]] }
const downloadExcelWorkbook = (sheets, filename) => {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, data }) => {
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, name.substring(0, 31));
  });
  XLSX.writeFile(wb, filename);
};

export default function AccountingApp() {
  const [darkMode, setDarkMode] = useState(false);
  const [activeTab, setActiveTab] = useState('dashboard');

  // Multi-currency State
  const [displayCurrency, setDisplayCurrency] = useState('VES');
  const [exchangeRate, setExchangeRate] = useState(36.50);
  const [showRateModal, setShowRateModal] = useState(false);

  // Persistent Application Data (valores por defecto mientras carga la base de datos)
  const [accounts, setAccounts] = useState(DEFAULT_ACCOUNTS);
  const [transactions, setTransactions] = useState([]);
  const [inventory, setInventory] = useState(DEFAULT_INVENTORY);

  // Controla si ya terminamos de leer la base de datos, para no sobreescribir
  // los datos guardados con los valores por defecto mientras carga.
  const [isDataLoaded, setIsDataLoaded] = useState(false);
  const [loadError, setLoadError] = useState(null);

  // Filtro de fechas para consultar los reportes (ej. 01/01/26 al 31/01/26)
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  // UI Modal States
  const [showResetModal, setShowResetModal] = useState(false);
  const [isResetting, setIsResetting] = useState(false);
  const [showAccountModal, setShowAccountModal] = useState(false);
  const [editingAccount, setEditingAccount] = useState(null);
  const [accCode, setAccCode] = useState('');
  const [accName, setAccName] = useState('');
  const [accType, setAccType] = useState('Activo');
  const [accCategory, setAccCategory] = useState('Cuenta');
  const [accNature, setAccNature] = useState('Deudora');
  const [accError, setAccError] = useState('');

  // Recarga los datos desde la nube: al abrir la app, cuando vuelves a esta pestaña del
  // navegador (no mientras estás escribiendo), o cuando le das al botón "Actualizar".
  const [pollTick, setPollTick] = useState(0);
  const [isRefreshing, setIsRefreshing] = useState(false);

  useEffect(() => {
    let cancelled = false;

    const loadAll = async () => {
      try {
        if (isDataLoaded) setIsRefreshing(true);
        const [acc, trx, inv, rate, dark] = await Promise.all([
          kvGet('procreas_accounts'),
          kvGet('procreas_transactions'),
          kvGet('procreas_inventory'),
          kvGet('procreas_rate'),
          kvGet('procreas_dark'),
        ]);

        if (cancelled) return;

        setAccounts(acc ?? DEFAULT_ACCOUNTS);
        setTransactions(trx ?? []);
        setInventory(inv ?? DEFAULT_INVENTORY);
        setExchangeRate(typeof rate === 'number' ? rate : 36.50);
        setDarkMode(!!dark);
        setLoadError(null);
      } catch (err) {
        if (!cancelled) {
          setLoadError('No se pudo conectar con la base de datos en la nube. Revisa tu conexión a internet o la configuración de Supabase (.env).');
        }
      } finally {
        if (!cancelled) {
          setIsDataLoaded(true);
          setIsRefreshing(false);
        }
      }
    };

    loadAll();
    return () => { cancelled = true; };
  }, [pollTick]);

  // Vuelve a consultar la nube cuando el usuario regresa a esta pestaña del navegador
  // (ej. estaba en otra app y vuelve). No interrumpe mientras se está escribiendo,
  // porque solo se dispara al cambiar de pestaña/ventana, no durante el tipeo.
  useEffect(() => {
    const onVisible = () => {
      if (document.visibilityState === 'visible') setPollTick(t => t + 1);
    };
    document.addEventListener('visibilitychange', onVisible);
    return () => document.removeEventListener('visibilitychange', onVisible);
  }, []);

  const refreshFromCloud = () => setPollTick(t => t + 1);

  // Guardar en la nube (solo después de que la carga inicial terminó,
  // para no pisar los datos guardados con los valores por defecto).
  useEffect(() => {
    if (!isDataLoaded) return;
    kvSet('procreas_accounts', accounts).catch(() => setLoadError('No se pudo guardar en la nube. Verifica tu conexión.'));
  }, [accounts, isDataLoaded]);

  useEffect(() => {
    if (!isDataLoaded) return;
    kvSet('procreas_transactions', transactions).catch(() => setLoadError('No se pudo guardar en la nube. Verifica tu conexión.'));
  }, [transactions, isDataLoaded]);

  useEffect(() => {
    if (!isDataLoaded) return;
    kvSet('procreas_inventory', inventory).catch(() => setLoadError('No se pudo guardar en la nube. Verifica tu conexión.'));
  }, [inventory, isDataLoaded]);

  useEffect(() => {
    if (!isDataLoaded) return;
    kvSet('procreas_rate', exchangeRate).catch(() => setLoadError('No se pudo guardar en la nube. Verifica tu conexión.'));
  }, [exchangeRate, isDataLoaded]);

  useEffect(() => {
    if (isDataLoaded) {
      kvSet('procreas_dark', darkMode).catch(() => {});
    }
    if (darkMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  }, [darkMode, isDataLoaded]);

  const executeResetBalances = () => {
    setIsResetting(true);
    setTimeout(() => {
      setTransactions([]);
      setIsResetting(false);
      setShowResetModal(false);
    }, 400);
  };

  const handleResetToDefaultAccounts = () => {
    if (window.confirm('¿Desea restaurar el plan de cuentas original de Grupo Procreas?')) {
      setAccounts(DEFAULT_ACCOUNTS);
    }
  };

  // Transacciones filtradas por el rango de fechas seleccionado (aplica a todos los reportes)
  const filteredTransactions = useMemo(() => {
    return transactions.filter(trx => {
      if (dateFrom && trx.date < dateFrom) return false;
      if (dateTo && trx.date > dateTo) return false;
      return true;
    });
  }, [transactions, dateFrom, dateTo]);

  const dateRangeLabel = () => {
    if (!dateFrom && !dateTo) return 'Todas las fechas';
    return `Del ${dateFrom || 'inicio'} al ${dateTo || 'día de hoy'}`;
  };

  // Convierte un monto (guardado en VES) a la moneda que se está mostrando en pantalla
  const convertAmt = (amount) => displayCurrency === 'USD' ? (amount / (exchangeRate || 1)) : amount;
  const todayStr = () => new Date().toISOString().split('T')[0];

  const buildJournalSheetData = () => {
    const rows = [
      ['LIBRO DIARIO - Grupo Procreas 929, C.A.'],
      [dateRangeLabel()],
      [`Moneda: ${displayCurrency}`],
      [],
      ['Fecha', 'Concepto / Descripción', 'Código Cuenta', 'Nombre Cuenta', `Debe (${displayCurrency})`, `Haber (${displayCurrency})`]
    ];
    filteredTransactions.forEach(trx => {
      trx.lines.forEach(line => {
        const account = accounts.find(a => a.id === line.accountId);
        rows.push([
          trx.date,
          trx.description,
          account ? account.code : 'N/A',
          account ? account.name : 'Cuenta Eliminada',
          line.type === 'Debe' ? convertAmt(line.amount) : '',
          line.type === 'Haber' ? convertAmt(line.amount) : ''
        ]);
      });
    });
    return rows;
  };

  const buildLedgerSheetData = (accData) => {
    const rows = [
      ['LIBRO MAYOR - Grupo Procreas 929, C.A.'],
      [`Cuenta: ${accData.code} - ${accData.name}`],
      [dateRangeLabel()],
      [`Moneda: ${displayCurrency}`],
      [],
      ['Fecha', 'Concepto', `Debe (${displayCurrency})`, `Haber (${displayCurrency})`]
    ];
    accData.entries.forEach(entry => {
      rows.push([
        entry.date,
        entry.description,
        entry.type === 'Debe' ? convertAmt(entry.amount) : '',
        entry.type === 'Haber' ? convertAmt(entry.amount) : ''
      ]);
    });
    rows.push([]);
    rows.push(['TOTALES', '', convertAmt(accData.debits), convertAmt(accData.credits)]);
    rows.push(['SALDO', '', '', convertAmt(accData.balance)]);
    return rows;
  };

  const buildTrialBalanceSheetData = () => {
    const activeAccounts = Object.values(accountBalances).filter(acc => acc.category === 'Cuenta' && (acc.debits > 0 || acc.credits > 0 || acc.balance !== 0));
    let totalDebits = 0, totalCredits = 0, totalDeudor = 0, totalAcreedor = 0;
    const rows = [
      ['BALANCE DE COMPROBACIÓN - Grupo Procreas 929, C.A.'],
      [dateRangeLabel()],
      [`Moneda: ${displayCurrency}`],
      [],
      ['Cuenta', `Debe (${displayCurrency})`, `Haber (${displayCurrency})`, `Saldo Deudor (${displayCurrency})`, `Saldo Acreedor (${displayCurrency})`]
    ];
    activeAccounts.forEach(acc => {
      totalDebits += acc.debits;
      totalCredits += acc.credits;
      if (acc.nature === 'Deudora') totalDeudor += acc.balance; else totalAcreedor += acc.balance;
      rows.push([
        `${acc.code} - ${acc.name}`,
        convertAmt(acc.debits),
        convertAmt(acc.credits),
        acc.nature === 'Deudora' && acc.balance > 0 ? convertAmt(acc.balance) : '',
        acc.nature === 'Acreedora' && acc.balance > 0 ? convertAmt(acc.balance) : ''
      ]);
    });
    rows.push(['TOTALES', convertAmt(totalDebits), convertAmt(totalCredits), convertAmt(totalDeudor), convertAmt(totalAcreedor)]);
    return rows;
  };

  const buildIncomeStatementSheetData = () => {
    const rows = [
      ['ESTADO DE RESULTADOS INTEGRAL - Grupo Procreas 929, C.A.'],
      [dateRangeLabel()],
      [`Moneda: ${displayCurrency}`],
      [],
      ['Concepto', `Monto (${displayCurrency})`],
      ['INGRESOS OPERATIVOS Y FINANCIEROS', '']
    ];
    Object.values(accountBalances).filter(a => a.type === 'Ingreso' && a.category === 'Cuenta').forEach(acc => {
      rows.push([`  ${acc.code} - ${acc.name}`, convertAmt(acc.balance)]);
    });
    rows.push(['Total Ingresos', convertAmt(financialStatements.revenue)]);
    rows.push([]);
    rows.push(['COSTOS Y GASTOS', '']);
    Object.values(accountBalances).filter(a => a.type === 'Gasto' && a.category === 'Cuenta').forEach(acc => {
      rows.push([`  ${acc.code} - ${acc.name}`, convertAmt(acc.balance)]);
    });
    rows.push(['Total Costos y Gastos', convertAmt(financialStatements.expenses)]);
    rows.push([]);
    rows.push(['UTILIDAD (PÉRDIDA) DEL EJERCICIO', convertAmt(financialStatements.netIncome)]);
    return rows;
  };

  const buildBalanceSheetSheetData = () => {
    const rows = [
      ['ESTADO DE SITUACIÓN FINANCIERA - Grupo Procreas 929, C.A.'],
      [`Corte al: ${dateTo || todayStr()}`],
      [`Moneda: ${displayCurrency}`],
      [],
      ['ACTIVO', '']
    ];
    Object.values(accountBalances).filter(a => a.type === 'Activo' && a.category === 'Cuenta').forEach(acc => {
      rows.push([`  ${acc.code} - ${acc.name}`, convertAmt(acc.balance)]);
    });
    rows.push(['Total Activo', convertAmt(financialStatements.assets)]);
    rows.push([]);
    rows.push(['PASIVO', '']);
    Object.values(accountBalances).filter(a => a.type === 'Pasivo' && a.category === 'Cuenta').forEach(acc => {
      rows.push([`  ${acc.code} - ${acc.name}`, convertAmt(acc.balance)]);
    });
    rows.push(['Total Pasivo', convertAmt(financialStatements.liabilities)]);
    rows.push([]);
    rows.push(['PATRIMONIO', '']);
    Object.values(accountBalances).filter(a => a.type === 'Patrimonio' && a.category === 'Cuenta').forEach(acc => {
      rows.push([`  ${acc.code} - ${acc.name}`, convertAmt(acc.balance)]);
    });
    rows.push(['Utilidad (Pérdida) del Ejercicio', convertAmt(financialStatements.netIncome)]);
    rows.push(['Total Patrimonio', convertAmt(financialStatements.equity)]);
    rows.push([]);
    rows.push(['TOTAL PASIVO + PATRIMONIO', convertAmt(financialStatements.liabilities + financialStatements.equity)]);
    return rows;
  };

  const buildInventorySheetData = () => {
    const rows = [
      ['LIBRO DE INVENTARIO - Grupo Procreas 929, C.A.'],
      [`Moneda: ${displayCurrency}`],
      [],
      ['Código Ref.', 'Artículo', 'Cantidad', `Costo Unitario (${displayCurrency})`, `Costo Total (${displayCurrency})`]
    ];
    if (inventory.length === 0) {
      rows.push(['(Sin existencias registradas)', '', '', '', '']);
    } else {
      inventory.forEach(item => {
        rows.push([item.code, item.name, item.quantity, convertAmt(item.unitCost), convertAmt(item.totalValue)]);
      });
    }
    const totalInventoryValue = inventory.reduce((sum, i) => sum + i.totalValue, 0);
    rows.push([]);
    rows.push(['Valorización Total', '', '', '', convertAmt(totalInventoryValue)]);
    return rows;
  };

  const exportJournalToExcel = () => {
    downloadExcelWorkbook([{ name: 'Libro Diario', data: buildJournalSheetData() }], `Libro_Diario_Procreas_${todayStr()}.xlsx`);
  };

  const exportLedgerToExcel = (accData) => {
    downloadExcelWorkbook([{ name: 'Libro Mayor', data: buildLedgerSheetData(accData) }], `Libro_Mayor_${accData.code}_Procreas_${todayStr()}.xlsx`);
  };

  const exportTrialBalanceToExcel = () => {
    downloadExcelWorkbook([{ name: 'B. Comprobación', data: buildTrialBalanceSheetData() }], `Balance_Comprobacion_Procreas_${todayStr()}.xlsx`);
  };

  const exportIncomeStatementToExcel = () => {
    downloadExcelWorkbook([{ name: 'E. de Resultados', data: buildIncomeStatementSheetData() }], `Estado_Resultados_Procreas_${todayStr()}.xlsx`);
  };

  const exportBalanceSheetToExcel = () => {
    downloadExcelWorkbook([{ name: 'E. de Situación', data: buildBalanceSheetSheetData() }], `Estado_Situacion_Procreas_${todayStr()}.xlsx`);
  };

  const exportInventoryToExcel = () => {
    downloadExcelWorkbook([{ name: 'Inventario', data: buildInventorySheetData() }], `Inventario_Procreas_${todayStr()}.xlsx`);
  };

  const exportAllReportsToExcel = () => {
    downloadExcelWorkbook([
      { name: 'Libro Diario', data: buildJournalSheetData() },
      { name: 'B. Comprobación', data: buildTrialBalanceSheetData() },
      { name: 'E. de Resultados', data: buildIncomeStatementSheetData() },
      { name: 'E. de Situación', data: buildBalanceSheetSheetData() },
      { name: 'Inventario', data: buildInventorySheetData() },
    ], `Reportes_Procreas_${todayStr()}.xlsx`);
  };

  // Barra reutilizable de filtro por fecha (Desde / Hasta), usada en cada reporte
  const DateRangeFilter = () => (
    <div className="flex flex-wrap items-end gap-3 bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
      <div>
        <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Desde</label>
        <input
          type="date"
          value={dateFrom}
          onChange={e => setDateFrom(e.target.value)}
          className="p-2 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none"
        />
      </div>
      <div>
        <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Hasta</label>
        <input
          type="date"
          value={dateTo}
          onChange={e => setDateTo(e.target.value)}
          className="p-2 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none"
        />
      </div>
      {(dateFrom || dateTo) && (
        <button
          onClick={() => { setDateFrom(''); setDateTo(''); }}
          className="text-xs text-slate-500 hover:text-red-500 underline pb-2.5"
        >
          Limpiar filtro
        </button>
      )}
      <span className="text-xs text-slate-400 pb-2.5 ml-auto">{filteredTransactions.length} asiento(s) en el periodo</span>
    </div>
  );

  const accountBalances = useMemo(() => {
    const balances = {};
    accounts.forEach(acc => {
      balances[acc.id] = { ...acc, debits: 0, credits: 0, balance: 0, entries: [] };
    });

    filteredTransactions.forEach(trx => {
      trx.lines.forEach(line => {
        if (balances[line.accountId]) {
          const amount = parseFloat(line.amount) || 0;
          balances[line.accountId].entries.push({
            date: trx.date,
            description: trx.description,
            type: line.type,
            amount: amount,
            trxId: trx.id
          });
          
          if (line.type === 'Debe') {
            balances[line.accountId].debits += amount;
          } else {
            balances[line.accountId].credits += amount;
          }
        }
      });
    });

    Object.values(balances).forEach(acc => {
      acc.balance = calculateBalance(acc.debits, acc.credits, acc.nature);
    });

    return balances;
  }, [filteredTransactions, accounts]);

  const financialStatements = useMemo(() => {
    let revenue = 0, expenses = 0, assets = 0, liabilities = 0, equity = 0;

    Object.values(accountBalances).forEach(acc => {
      if (acc.category === 'Cuenta') {
        if (acc.type === 'Ingreso') revenue += acc.balance;
        if (acc.type === 'Gasto') expenses += acc.balance;
        if (acc.type === 'Activo') assets += acc.balance;
        if (acc.type === 'Pasivo') liabilities += acc.balance;
        if (acc.type === 'Patrimonio') equity += acc.balance;
      }
    });

    const netIncome = revenue - expenses;
    const totalEquity = equity + netIncome;

    return {
      revenue, expenses, netIncome, assets, liabilities, equity: totalEquity,
      isBalanced: Math.abs(assets - (liabilities + totalEquity)) < 0.01
    };
  }, [accountBalances]);

  // Account Management Handlers
  const handleOpenAddAccount = () => {
    setEditingAccount(null);
    setAccCode('');
    setAccName('');
    setAccType('Activo');
    setAccCategory('Cuenta');
    setAccNature('Deudora');
    setAccError('');
    setShowAccountModal(true);
  };

  const handleOpenEditAccount = (acc) => {
    setEditingAccount(acc);
    setAccCode(acc.code);
    setAccName(acc.name);
    setAccType(acc.type);
    setAccCategory(acc.category);
    setAccNature(acc.nature);
    setAccError('');
    setShowAccountModal(true);
  };

  const handleSaveAccount = () => {
    if (!accCode.trim() || !accName.trim()) {
      setAccError('El código y el nombre son obligatorios.');
      return;
    }

    if (editingAccount) {
      setAccounts(accounts.map(a => a.id === editingAccount.id ? {
        ...a,
        code: accCode.trim(),
        name: accName.trim(),
        type: accType,
        category: accCategory,
        nature: accNature
      } : a));
    } else {
      const newAcc = {
        id: 'acc_' + Date.now(),
        code: accCode.trim(),
        name: accName.trim(),
        type: accType,
        category: accCategory,
        nature: accNature
      };
      setAccounts([...accounts, newAcc].sort((a,b) => a.code.localeCompare(b.code)));
    }
    setShowAccountModal(false);
  };

  const renderNav = () => {
    const navItems = [
      { id: 'dashboard', icon: LayoutDashboard, label: 'Dashboard' },
      { id: 'cuentas', icon: Book, label: 'Plan de Cuentas' },
      { id: 'diario', icon: BookOpen, label: 'Libro Diario' },
      { id: 'mayor', icon: BookOpen, label: 'Libro Mayor' },
      { id: 'inventario', icon: Package, label: 'Inventario' },
      { id: 'comprobacion', icon: FileText, label: 'B. Comprobación' },
      { id: 'resultados', icon: BarChart2, label: 'E. de Resultados' },
      { id: 'situacion', icon: PieChart, label: 'E. de Situación' }
    ];

    return (
      <div className="w-64 bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800 h-screen flex flex-col transition-colors duration-200 fixed lg:static z-20 shadow-sm">
        <div className="p-5 border-b border-slate-100 dark:border-slate-800">
          <h1 className="text-lg font-extrabold text-slate-800 dark:text-white flex items-center gap-2">
            <span className="text-emerald-600 dark:text-emerald-400">PROCREAS</span>
          </h1>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">Contabilidad NIIF</p>
        </div>

        {/* Currency Switcher & Rate Widget */}
        <div className="px-4 py-3 bg-slate-50 dark:bg-slate-800/50 border-b border-slate-200 dark:border-slate-800">
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-semibold text-slate-500 dark:text-slate-400">Moneda</span>
            <div className="flex bg-slate-200 dark:bg-slate-700 p-0.5 rounded-lg">
              <button 
                onClick={() => setDisplayCurrency('VES')}
                className={`px-2.5 py-1 text-xs font-bold rounded-md transition-all ${displayCurrency === 'VES' ? 'bg-white dark:bg-slate-800 text-slate-800 dark:text-white shadow-sm' : 'text-slate-600 dark:text-slate-400'}`}
              >
                VES
              </button>
              <button 
                onClick={() => setDisplayCurrency('USD')}
                className={`px-2.5 py-1 text-xs font-bold rounded-md transition-all ${displayCurrency === 'USD' ? 'bg-white dark:bg-slate-800 text-slate-800 dark:text-white shadow-sm' : 'text-slate-600 dark:text-slate-400'}`}
              >
                USD
              </button>
            </div>
          </div>
          <div className="flex items-center justify-between text-xs text-slate-500 dark:text-slate-400 pt-1">
            <span>Tasa: {exchangeRate.toFixed(2)} VES</span>
            <button 
              onClick={() => setShowRateModal(true)}
              className="text-emerald-600 dark:text-emerald-400 font-medium hover:underline"
            >
              Cambiar
            </button>
          </div>
        </div>

        <nav className="flex-1 px-3 py-3 space-y-1 overflow-y-auto">
          {navItems.map(item => {
            const Icon = item.icon;
            return (
              <button
                key={item.id}
                onClick={() => setActiveTab(item.id)}
                className={`w-full flex items-center gap-3 px-3.5 py-2.5 rounded-xl transition-all duration-200 text-sm ${
                  activeTab === item.id 
                    ? 'bg-emerald-50 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 font-semibold shadow-sm' 
                    : 'text-slate-600 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-800/50'
                }`}
              >
                <Icon size={18} />
                <span>{item.label}</span>
              </button>
            );
          })}
        </nav>

        <div className="p-4 border-t border-slate-200 dark:border-slate-800">
          <div className="flex items-center gap-2 mb-3 text-xs text-slate-500 dark:text-slate-400 justify-center">
             <Database size={13} className={loadError ? 'text-red-500' : 'text-emerald-500'} />
             {loadError ? 'Sin conexión a la nube' : 'Base de Datos en la Nube'}
             <div className={`w-2 h-2 rounded-full ${loadError ? 'bg-red-500' : 'bg-emerald-500'}`}></div>
          </div>
          <button
            onClick={refreshFromCloud}
            disabled={isRefreshing}
            className="w-full flex items-center justify-center gap-2 px-3 py-1.5 mb-2 rounded-lg text-xs text-slate-500 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors disabled:opacity-50"
          >
            <RefreshCw size={12} className={isRefreshing ? 'animate-spin' : ''} />
            {isRefreshing ? 'Actualizando...' : 'Actualizar datos'}
          </button>
          <button
            onClick={() => setDarkMode(!darkMode)}
            className="w-full flex items-center justify-center gap-2 px-3 py-2 rounded-lg bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700 transition-colors text-sm"
          >
            {darkMode ? <Sun size={16} /> : <Moon size={16} />}
            {darkMode ? 'Modo Claro' : 'Modo Oscuro'}
          </button>
          <button
            onClick={() => supabase.auth.signOut()}
            className="w-full flex items-center justify-center gap-2 px-3 py-2 mt-2 rounded-lg text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 transition-colors text-xs"
          >
            Cerrar sesión
          </button>
        </div>
      </div>
    );
  };

  const AccountModal = () => {
    if (!showAccountModal) return null;
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm animate-in fade-in duration-200 p-4">
        <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-xl w-full max-w-md overflow-hidden border border-slate-200 dark:border-slate-700 p-6">
          <div className="flex justify-between items-center mb-4">
            <h3 className="text-lg font-bold text-slate-800 dark:text-white">
              {editingAccount ? 'Modificar Cuenta' : 'Nueva Cuenta Contable'}
            </h3>
            <button onClick={() => setShowAccountModal(false)} className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200">
              <X size={20} />
            </button>
          </div>

          <div className="space-y-4">
            <div>
              <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Código (Ej. 1.1.02.002)</label>
              <input 
                type="text" 
                value={accCode} 
                onChange={e => setAccCode(e.target.value)}
                placeholder="1.1.02.002"
                className="w-full p-2.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white font-mono text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
              />
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Nombre de la Cuenta</label>
              <input 
                type="text" 
                value={accName} 
                onChange={e => setAccName(e.target.value)}
                placeholder="Ej. BANCO PROVINCIAL"
                className="w-full p-2.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
              />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Tipo</label>
                <select 
                  value={accType} 
                  onChange={e => setAccType(e.target.value)}
                  className="w-full p-2.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                >
                  <option value="Activo">Activo</option>
                  <option value="Pasivo">Pasivo</option>
                  <option value="Patrimonio">Patrimonio</option>
                  <option value="Ingreso">Ingreso</option>
                  <option value="Gasto">Gasto</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Naturaleza</label>
                <select 
                  value={accNature} 
                  onChange={e => setAccNature(e.target.value)}
                  className="w-full p-2.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
                >
                  <option value="Deudora">Deudora</option>
                  <option value="Acreedora">Acreedora</option>
                </select>
              </div>
            </div>
            <div>
              <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">Categoría / Nivel</label>
              <select 
                value={accCategory} 
                onChange={e => setAccCategory(e.target.value)}
                className="w-full p-2.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white text-sm focus:ring-2 focus:ring-emerald-500 outline-none"
              >
                <option value="Cuenta">Cuenta (Movimiento)</option>
                <option value="Grupo">Grupo</option>
                <option value="Titulo">Título</option>
              </select>
            </div>

            {accError && <p className="text-red-500 text-xs">{accError}</p>}
          </div>

          <div className="flex gap-3 justify-end mt-6">
            <button 
              onClick={() => setShowAccountModal(false)}
              className="px-4 py-2 rounded-lg font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors text-sm"
            >
              Cancelar
            </button>
            <button 
              onClick={handleSaveAccount}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg font-medium transition-colors shadow-sm text-sm flex items-center gap-1.5"
            >
              <Save size={16} /> Guardar
            </button>
          </div>
        </div>
      </div>
    );
  };

  const ResetConfirmModal = () => {
    if (!showResetModal) return null;
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm animate-in fade-in duration-200 p-4">
        <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-xl w-full max-w-md overflow-hidden border border-slate-200 dark:border-slate-700 p-6">
          <div className="w-12 h-12 rounded-full bg-red-100 dark:bg-red-900/30 flex items-center justify-center mb-4 text-red-600 dark:text-red-400">
            <AlertTriangle size={24} />
          </div>
          <h3 className="text-xl font-bold text-slate-800 dark:text-white mb-2">¿Reiniciar todos los saldos a 0?</h3>
          <p className="text-slate-600 dark:text-slate-400 mb-6 text-sm">
            Se eliminarán todos los asientos del Libro Diario, restableciendo los saldos a cero. Tu plan de cuentas se mantendrá intacto.
          </p>
          <div className="flex gap-3 justify-end">
            <button 
              onClick={() => setShowResetModal(false)}
              disabled={isResetting}
              className="px-4 py-2 rounded-lg font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors text-sm"
            >
              Cancelar
            </button>
            <button 
              onClick={executeResetBalances}
              disabled={isResetting}
              className="flex items-center gap-2 bg-red-600 hover:bg-red-700 text-white px-4 py-2 rounded-lg font-medium transition-colors shadow-sm text-sm disabled:opacity-70"
            >
              {isResetting ? <Loader2 size={16} className="animate-spin" /> : <Trash2 size={16} />}
              Sí, Reiniciar
            </button>
          </div>
        </div>
      </div>
    );
  };

  const ExchangeRateModal = () => {
    if (!showRateModal) return null;
    const [tempRate, setTempRate] = useState(exchangeRate);
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 backdrop-blur-sm animate-in fade-in duration-200 p-4">
        <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-xl w-full max-w-sm overflow-hidden border border-slate-200 dark:border-slate-700 p-6">
          <h3 className="text-lg font-bold text-slate-800 dark:text-white mb-2">Tasa de Cambio (VES / USD)</h3>
          <p className="text-sm text-slate-600 dark:text-slate-400 mb-4">
            Establece la tasa oficial o paralela actual para convertir reportes y registrar operaciones en dólares.
          </p>
          <div className="mb-6">
            <label className="block text-xs font-medium text-slate-500 dark:text-slate-400 mb-1">1 USD en VES</label>
            <input 
              type="number" 
              step="0.01" 
              value={tempRate} 
              onChange={e => setTempRate(e.target.value)}
              className="w-full p-2.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white font-mono text-lg focus:ring-2 focus:ring-emerald-500 outline-none"
            />
          </div>
          <div className="flex gap-3 justify-end">
            <button 
              onClick={() => setShowRateModal(false)}
              className="px-4 py-2 rounded-lg font-medium text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors text-sm"
            >
              Cancelar
            </button>
            <button 
              onClick={() => {
                if (parseFloat(tempRate) > 0) {
                  setExchangeRate(parseFloat(tempRate));
                }
                setShowRateModal(false);
              }}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg font-medium transition-colors shadow-sm text-sm"
            >
              Guardar Tasa
            </button>
          </div>
        </div>
      </div>
    );
  };

  const ViewDashboard = () => (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 dark:text-white">Grupo Procreas 929, C.A.</h2>
          <p className="text-sm text-slate-500 dark:text-slate-400">Sistema Contable Adaptado a Normas NIIF ({displayCurrency})</p>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-slate-500 dark:text-slate-400 text-sm font-medium mb-2">Total Activos</h3>
          <p className="text-2xl font-bold text-slate-800 dark:text-white">{formatCurrency(financialStatements.assets, displayCurrency, exchangeRate)}</p>
        </div>
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-slate-500 dark:text-slate-400 text-sm font-medium mb-2">Total Pasivos</h3>
          <p className="text-2xl font-bold text-red-500 dark:text-red-400">{formatCurrency(financialStatements.liabilities, displayCurrency, exchangeRate)}</p>
        </div>
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-slate-500 dark:text-slate-400 text-sm font-medium mb-2">Patrimonio Total</h3>
          <p className="text-2xl font-bold text-blue-500 dark:text-blue-400">{formatCurrency(financialStatements.equity, displayCurrency, exchangeRate)}</p>
        </div>
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-slate-500 dark:text-slate-400 text-sm font-medium mb-2">Utilidad Neta</h3>
          <p className={`text-2xl font-bold ${financialStatements.netIncome >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-500'}`}>
            {formatCurrency(financialStatements.netIncome, displayCurrency, exchangeRate)}
          </p>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-lg font-bold text-slate-800 dark:text-white mb-4">Ecuación Contable NIIF</h3>
          <div className="flex flex-wrap items-center justify-between p-4 bg-slate-50 dark:bg-slate-900 rounded-xl gap-2">
            <div className="text-center flex-1 min-w-[100px]">
              <p className="text-xs text-slate-500 dark:text-slate-400 mb-1">Activos</p>
              <p className="text-lg font-bold text-slate-800 dark:text-white">{formatCurrency(financialStatements.assets, displayCurrency, exchangeRate)}</p>
            </div>
            <div className="text-xl font-bold text-slate-300 dark:text-slate-600">=</div>
            <div className="text-center flex-1 min-w-[100px]">
              <p className="text-xs text-slate-500 dark:text-slate-400 mb-1">Pasivos</p>
              <p className="text-lg font-bold text-slate-800 dark:text-white">{formatCurrency(financialStatements.liabilities, displayCurrency, exchangeRate)}</p>
            </div>
            <div className="text-xl font-bold text-slate-300 dark:text-slate-600">+</div>
            <div className="text-center flex-1 min-w-[100px]">
              <p className="text-xs text-slate-500 dark:text-slate-400 mb-1">Patrimonio</p>
              <p className="text-lg font-bold text-slate-800 dark:text-white">{formatCurrency(financialStatements.equity, displayCurrency, exchangeRate)}</p>
            </div>
          </div>
          {financialStatements.isBalanced ? (
            <div className="mt-4 flex items-center text-emerald-600 dark:text-emerald-400 text-sm font-medium">
              <CheckCircle size={16} className="mr-2" /> Ecuación contable cuadrada correctamente.
            </div>
          ) : (
            <div className="mt-4 flex items-center text-red-500 text-sm font-medium">
              <AlertCircle size={16} className="mr-2" /> La ecuación presenta diferencias en los saldos.
            </div>
          )}
        </div>

        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="text-lg font-bold text-slate-800 dark:text-white mb-4">Herramientas</h3>
          <div className="space-y-3">
            <button 
              onClick={exportAllReportsToExcel}
              className="w-full flex items-center justify-between p-3 rounded-xl bg-slate-50 hover:bg-slate-100 dark:bg-slate-900 dark:hover:bg-slate-800 border border-slate-200 dark:border-slate-700 transition-colors text-sm group"
            >
              <div className="flex items-center gap-3 text-slate-700 dark:text-slate-300 font-medium">
                <FileSpreadsheet size={18} className="text-emerald-600" />
                Exportar todos los reportes a Excel
              </div>
              <DownloadCloud size={16} className="text-slate-400 group-hover:text-emerald-600 transition-colors" />
            </button>

            <button 
              onClick={() => setShowResetModal(true)}
              className="w-full flex items-center justify-between p-3 rounded-xl bg-slate-50 hover:bg-red-50 dark:bg-slate-900 dark:hover:bg-red-900/20 border border-slate-200 dark:border-slate-700 transition-colors text-sm group"
            >
              <div className="flex items-center gap-3 text-slate-700 dark:text-slate-300 font-medium group-hover:text-red-600 dark:group-hover:text-red-400 transition-colors">
                <RefreshCw size={18} className="text-red-500" />
                Reiniciar saldos a 0
              </div>
            </button>
          </div>
        </div>
      </div>
    </div>
  );

  const ViewChartOfAccounts = () => (
    <div className="space-y-6 animate-in fade-in duration-300">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 dark:text-white">Plan de Cuentas</h2>
          <p className="text-xs text-slate-500 dark:text-slate-400">Grupo Procreas 929, C.A.</p>
        </div>
        <div className="flex gap-2">
          <button 
            onClick={handleResetToDefaultAccounts}
            className="flex items-center gap-1.5 text-xs bg-slate-200 hover:bg-slate-300 dark:bg-slate-700 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 px-3 py-2 rounded-xl transition-colors font-medium"
          >
            Restaurar Plan Original
          </button>
          <button 
            onClick={handleOpenAddAccount}
            className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white font-medium px-4 py-2 rounded-xl transition-colors shadow-sm text-sm"
          >
            <Plus size={16} /> Nueva Cuenta
          </button>
        </div>
      </div>

      <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50 dark:bg-slate-900 border-b border-slate-200 dark:border-slate-700">
                <th className="p-4 text-sm font-semibold text-slate-600 dark:text-slate-300">Código</th>
                <th className="p-4 text-sm font-semibold text-slate-600 dark:text-slate-300">Nombre de la Cuenta</th>
                <th className="p-4 text-sm font-semibold text-slate-600 dark:text-slate-300">Tipo</th>
                <th className="p-4 text-sm font-semibold text-slate-600 dark:text-slate-300">Naturaleza</th>
                <th className="p-4 text-sm font-semibold text-slate-600 dark:text-slate-300">Nivel</th>
                <th className="p-4 text-sm font-semibold text-slate-600 dark:text-slate-300 text-right">Acciones</th>
              </tr>
            </thead>
            <tbody>
              {accounts.map(acc => (
                <tr key={acc.id} className="border-b border-slate-100 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors">
                  <td className="p-4 text-slate-800 dark:text-slate-200 font-mono text-xs">{acc.code}</td>
                  <td className={`p-4 text-slate-800 dark:text-slate-200 text-sm ${acc.category === 'Titulo' ? 'font-bold' : acc.category === 'Grupo' ? 'font-semibold pl-6' : 'pl-10'}`}>
                    {acc.name}
                  </td>
                  <td className="p-4 text-slate-600 dark:text-slate-400 text-xs">{acc.type}</td>
                  <td className="p-4 text-slate-600 dark:text-slate-400 text-xs">
                    <span className={`px-2 py-0.5 rounded-full text-xs ${acc.nature === 'Deudora' ? 'bg-blue-100 text-blue-700 dark:bg-blue-900/30 dark:text-blue-300' : 'bg-purple-100 text-purple-700 dark:bg-purple-900/30 dark:text-purple-300'}`}>
                      {acc.nature}
                    </span>
                  </td>
                  <td className="p-4 text-slate-600 dark:text-slate-400 text-xs">{acc.category}</td>
                  <td className="p-4 text-right">
                    <button 
                      onClick={() => handleOpenEditAccount(acc)}
                      className="flex items-center gap-1 text-xs bg-slate-100 dark:bg-slate-700 hover:bg-slate-200 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 px-3 py-1.5 rounded-lg font-medium transition-colors ml-auto"
                    >
                      <Edit3 size={14} /> Editar
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );

  const ViewJournal = () => {
    const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
    const [description, setDescription] = useState('');
    const [trxCurrency, setTrxCurrency] = useState('VES');
    const [lines, setLines] = useState([{ accountId: '', type: 'Debe', amount: '' }, { accountId: '', type: 'Haber', amount: '' }]);
    const [error, setError] = useState('');
    const [isSaving, setIsSaving] = useState(false);

    const selectableAccounts = accounts.filter(a => a.category === 'Cuenta');

    const handleAddLine = () => setLines([...lines, { accountId: '', type: 'Debe', amount: '' }]);
    const handleRemoveLine = (index) => setLines(lines.filter((_, i) => i !== index));
    const handleLineChange = (index, field, value) => {
      const newLines = [...lines];
      newLines[index][field] = value;
      setLines(newLines);
    };

    const handleSaveTransaction = () => {
      setError('');
      if (!description) {
        setError('Debe incluir una descripción.');
        return;
      }

      let totalDebits = 0;
      let totalCredits = 0;
      
      for (const line of lines) {
        if (!line.accountId || !line.amount || isNaN(line.amount) || Number(line.amount) <= 0) {
          setError('Todas las líneas deben tener una cuenta y un monto válido mayor a 0.');
          return;
        }
        const val = Number(line.amount);
        if (line.type === 'Debe') totalDebits += val;
        else totalCredits += val;
      }

      if (Math.abs(totalDebits - totalCredits) > 0.01) {
        setError(`El asiento no cuadra. Debe: ${formatCurrency(totalDebits, trxCurrency, exchangeRate)} | Haber: ${formatCurrency(totalCredits, trxCurrency, exchangeRate)}`);
        return;
      }

      setIsSaving(true);
      const conversionMultiplier = trxCurrency === 'USD' ? exchangeRate : 1;

      const newTrx = {
        id: 'trx_' + Date.now().toString(),
        date,
        description: `${description} (${trxCurrency})`,
        lines: lines.map(l => ({ 
          ...l, 
          amount: Number(l.amount) * conversionMultiplier 
        }))
      };

      setTransactions([newTrx, ...transactions]);
      setIsSaving(false);
      setDescription('');
      setLines([{ accountId: '', type: 'Debe', amount: '' }, { accountId: '', type: 'Haber', amount: '' }]);
    };

    return (
      <div className="space-y-6 animate-in fade-in duration-300">
        <div className="flex justify-between items-center">
          <h2 className="text-2xl font-bold text-slate-800 dark:text-white">Libro Diario</h2>
          <button onClick={exportJournalToExcel} className="flex items-center gap-2 text-xs bg-emerald-100 hover:bg-emerald-200 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300 px-3.5 py-2 rounded-xl transition-colors font-medium">
            <FileSpreadsheet size={16} /> Exportar Excel
          </button>
        </div>

        <DateRangeFilter />
        
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <div className="flex justify-between items-center mb-4">
            <h3 className="text-base font-semibold text-slate-800 dark:text-white">Registrar Nuevo Asiento</h3>
            <div className="flex items-center gap-2">
              <span className="text-xs text-slate-500">Moneda:</span>
              <div className="flex bg-slate-100 dark:bg-slate-900 p-1 rounded-lg border border-slate-200 dark:border-slate-700">
                <button 
                  onClick={() => setTrxCurrency('VES')}
                  className={`px-3 py-1 text-xs font-bold rounded-md transition-all ${trxCurrency === 'VES' ? 'bg-emerald-600 text-white shadow-sm' : 'text-slate-600 dark:text-slate-400'}`}
                >
                  VES
                </button>
                <button 
                  onClick={() => setTrxCurrency('USD')}
                  className={`px-3 py-1 text-xs font-bold rounded-md transition-all ${trxCurrency === 'USD' ? 'bg-emerald-600 text-white shadow-sm' : 'text-slate-600 dark:text-slate-400'}`}
                >
                  USD
                </button>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
            <div>
              <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Fecha</label>
              <input type="date" value={date} onChange={e => setDate(e.target.value)} className="w-full p-2.5 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none" />
            </div>
            <div className="md:col-span-2">
              <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Concepto / Descripción</label>
              <input type="text" placeholder="Ej. Pago de mantenimiento de ascensores..." value={description} onChange={e => setDescription(e.target.value)} className="w-full p-2.5 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none" />
            </div>
          </div>

          <div className="space-y-3 mb-4">
            {lines.map((line, index) => (
              <div key={index} className="flex flex-wrap md:flex-nowrap gap-3 items-end">
                <div className="flex-1 min-w-[200px]">
                  <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Cuenta</label>
                  <select value={line.accountId} onChange={e => handleLineChange(index, 'accountId', e.target.value)} className="w-full p-2.5 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none">
                    <option value="">Seleccione cuenta...</option>
                    {selectableAccounts.map(a => <option key={a.id} value={a.id}>{a.code} - {a.name}</option>)}
                  </select>
                </div>
                <div className="w-32">
                  <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Tipo</label>
                  <select value={line.type} onChange={e => handleLineChange(index, 'type', e.target.value)} className="w-full p-2.5 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none">
                    <option value="Debe">Debe</option>
                    <option value="Haber">Haber</option>
                  </select>
                </div>
                <div className="w-40">
                  <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-1">Monto ({trxCurrency})</label>
                  <input type="number" step="0.01" min="0" placeholder="0.00" value={line.amount} onChange={e => handleLineChange(index, 'amount', e.target.value)} className="w-full p-2.5 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none" />
                </div>
                {lines.length > 2 && (
                  <button onClick={() => handleRemoveLine(index)} className="p-2.5 text-red-500 hover:bg-red-50 dark:hover:bg-red-900/20 rounded-lg transition-colors">
                    <Trash2 size={18} />
                  </button>
                )}
              </div>
            ))}
          </div>

          {error && <p className="text-red-500 text-xs mb-4">{error}</p>}

          <div className="flex justify-between items-center border-t border-slate-100 dark:border-slate-700 pt-4">
            <button onClick={handleAddLine} className="flex items-center gap-2 text-emerald-600 dark:text-emerald-400 text-xs font-semibold hover:bg-emerald-50 dark:hover:bg-emerald-900/20 px-3 py-2 rounded-lg transition-colors">
              <Plus size={16} /> Agregar línea
            </button>
            <button disabled={isSaving} onClick={handleSaveTransaction} className="flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-emerald-400 text-white text-sm font-medium px-5 py-2.5 rounded-xl transition-colors shadow-sm">
              {isSaving ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />} Registrar Asiento
            </button>
          </div>
        </div>

        <div className="space-y-4">
          <h3 className="text-base font-semibold text-slate-800 dark:text-white">Historial de Asientos</h3>
          {filteredTransactions.length === 0 ? (
            <p className="text-slate-500 text-xs italic">No hay asientos registrados en el periodo seleccionado.</p>
          ) : (
            filteredTransactions.map((trx, idx) => (
              <div key={trx.id} className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 overflow-hidden">
                <div className="bg-slate-50 dark:bg-slate-900 p-4 border-b border-slate-100 dark:border-slate-700 flex justify-between items-center">
                  <div>
                    <span className="font-bold text-sm text-slate-800 dark:text-white mr-3">Asiento #{filteredTransactions.length - idx}</span>
                    <span className="text-xs text-slate-500">{trx.date}</span>
                  </div>
                  <span className="text-xs font-medium text-slate-700 dark:text-slate-300">{trx.description}</span>
                </div>
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-100 dark:border-slate-700 text-slate-400">
                      <th className="p-3">Cuenta</th>
                      <th className="p-3 text-right">Debe</th>
                      <th className="p-3 text-right">Haber</th>
                    </tr>
                  </thead>
                  <tbody>
                    {trx.lines.map((line, i) => {
                      const account = accounts.find(a => a.id === line.accountId);
                      return (
                        <tr key={i} className="border-b border-slate-50 dark:border-slate-800/50">
                          <td className={`p-3 text-slate-800 dark:text-slate-200 ${line.type === 'Haber' ? 'pl-6' : ''}`}>
                            {account ? `${account.code} - ${account.name}` : 'Cuenta Eliminada'}
                          </td>
                          <td className="p-3 text-right font-mono text-slate-800 dark:text-slate-200">
                            {line.type === 'Debe' ? formatCurrency(line.amount, displayCurrency, exchangeRate) : ''}
                          </td>
                          <td className="p-3 text-right font-mono text-slate-800 dark:text-slate-200">
                            {line.type === 'Haber' ? formatCurrency(line.amount, displayCurrency, exchangeRate) : ''}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            ))
          )}
        </div>
      </div>
    );
  };

  const ViewLedger = () => {
    const [selectedAccount, setSelectedAccount] = useState('');
    const accountList = accounts.filter(a => a.category === 'Cuenta');
    const accData = selectedAccount ? accountBalances[selectedAccount] : null;

    return (
      <div className="space-y-6 animate-in fade-in duration-300">
        <h2 className="text-2xl font-bold text-slate-800 dark:text-white">Libro Mayor</h2>

        <DateRangeFilter />

        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <label className="block text-xs font-medium text-slate-600 dark:text-slate-400 mb-2">Seleccione una cuenta contable</label>
          <select 
            value={selectedAccount} 
            onChange={e => setSelectedAccount(e.target.value)} 
            className="w-full max-w-md p-2.5 text-sm rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:ring-2 focus:ring-emerald-500 outline-none"
          >
            <option value="">-- Seleccionar Cuenta --</option>
            {accountList.map(a => <option key={a.id} value={a.id}>{a.code} - {a.name}</option>)}
          </select>
        </div>

        {accData && (
          <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 overflow-hidden">
            <div className="p-6 border-b border-slate-100 dark:border-slate-700 flex flex-col sm:flex-row justify-between items-start sm:items-center bg-slate-50 dark:bg-slate-900 gap-2">
              <div>
                <h3 className="text-lg font-bold text-slate-800 dark:text-white">{accData.code} - {accData.name}</h3>
                <p className="text-xs text-slate-500">Naturaleza: {accData.nature}</p>
              </div>
              <div className="flex items-center gap-4">
                <div className="text-right">
                  <p className="text-xs text-slate-500">Saldo Actual ({displayCurrency})</p>
                  <p className={`text-xl font-bold ${accData.balance >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-500'}`}>
                    {formatCurrency(accData.balance, displayCurrency, exchangeRate)}
                  </p>
                </div>
                <button onClick={() => exportLedgerToExcel(accData)} className="flex items-center gap-2 text-xs bg-emerald-100 hover:bg-emerald-200 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300 px-3.5 py-2 rounded-xl transition-colors font-medium whitespace-nowrap">
                  <FileSpreadsheet size={16} /> Exportar Excel
                </button>
              </div>
            </div>
            
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 dark:bg-slate-900/50 border-b border-slate-200 dark:border-slate-700">
                    <th className="p-4 font-semibold text-slate-600 dark:text-slate-300">Fecha</th>
                    <th className="p-4 font-semibold text-slate-600 dark:text-slate-300">Concepto</th>
                    <th className="p-4 font-semibold text-slate-600 dark:text-slate-300 text-right">Debe</th>
                    <th className="p-4 font-semibold text-slate-600 dark:text-slate-300 text-right">Haber</th>
                  </tr>
                </thead>
                <tbody>
                  {accData.entries.length === 0 ? (
                    <tr>
                      <td colSpan="4" className="p-4 text-center text-slate-500 italic">No hay movimientos para esta cuenta.</td>
                    </tr>
                  ) : (
                    accData.entries.map((entry, idx) => (
                      <tr key={idx} className="border-b border-slate-100 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800/50">
                        <td className="p-4 text-slate-800 dark:text-slate-200">{entry.date}</td>
                        <td className="p-4 text-slate-800 dark:text-slate-200">{entry.description}</td>
                        <td className="p-4 text-right font-mono text-slate-800 dark:text-slate-200">
                          {entry.type === 'Debe' ? formatCurrency(entry.amount, displayCurrency, exchangeRate) : '-'}
                        </td>
                        <td className="p-4 text-right font-mono text-slate-800 dark:text-slate-200">
                          {entry.type === 'Haber' ? formatCurrency(entry.amount, displayCurrency, exchangeRate) : '-'}
                        </td>
                      </tr>
                    ))
                  )}
                  {accData.entries.length > 0 && (
                    <tr className="bg-slate-50 dark:bg-slate-900 font-bold">
                      <td colSpan="2" className="p-4 text-right text-slate-800 dark:text-slate-200">TOTALES:</td>
                      <td className="p-4 text-right font-mono text-slate-800 dark:text-slate-200">{formatCurrency(accData.debits, displayCurrency, exchangeRate)}</td>
                      <td className="p-4 text-right font-mono text-slate-800 dark:text-slate-200">{formatCurrency(accData.credits, displayCurrency, exchangeRate)}</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    );
  };

  const ViewTrialBalance = () => {
    const activeAccounts = Object.values(accountBalances).filter(acc => acc.category === 'Cuenta' && (acc.debits > 0 || acc.credits > 0 || acc.balance !== 0));
    
    let totalDebits = 0, totalCredits = 0, totalDebtorBalance = 0, totalCreditorBalance = 0;
    activeAccounts.forEach(acc => {
      totalDebits += acc.debits;
      totalCredits += acc.credits;
      if (acc.nature === 'Deudora') totalDebtorBalance += acc.balance;
      else totalCreditorBalance += acc.balance;
    });

    const isSumBalanced = Math.abs(totalDebits - totalCredits) < 0.01;
    const isBalBalanced = Math.abs(totalDebtorBalance - totalCreditorBalance) < 0.01;

    return (
      <div className="space-y-6 animate-in fade-in duration-300">
        <div className="flex justify-between items-center">
          <h2 className="text-2xl font-bold text-slate-800 dark:text-white">Balance de Comprobación</h2>
          <button onClick={exportTrialBalanceToExcel} className="flex items-center gap-2 text-xs bg-emerald-100 hover:bg-emerald-200 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300 px-3.5 py-2 rounded-xl transition-colors font-medium">
            <FileSpreadsheet size={16} /> Exportar Excel
          </button>
        </div>

        <DateRangeFilter />

        <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 dark:bg-slate-900 border-b border-slate-200 dark:border-slate-700">
                  <th rowSpan="2" className="p-4 font-semibold text-slate-600 dark:text-slate-300 align-bottom border-r dark:border-slate-700">Cuenta</th>
                  <th colSpan="2" className="p-2 font-semibold text-center text-slate-600 dark:text-slate-300 border-b border-r dark:border-slate-700">Sumas</th>
                  <th colSpan="2" className="p-2 font-semibold text-center text-slate-600 dark:text-slate-300 border-b dark:border-slate-700">Saldos</th>
                </tr>
                <tr className="bg-slate-50 dark:bg-slate-900 border-b border-slate-200 dark:border-slate-700">
                  <th className="p-2 font-semibold text-right text-slate-600 dark:text-slate-300 border-r dark:border-slate-700">Debe</th>
                  <th className="p-2 font-semibold text-right text-slate-600 dark:text-slate-300 border-r dark:border-slate-700">Haber</th>
                  <th className="p-2 font-semibold text-right text-slate-600 dark:text-slate-300 border-r dark:border-slate-700">Deudor</th>
                  <th className="p-2 font-semibold text-right text-slate-600 dark:text-slate-300">Acreedor</th>
                </tr>
              </thead>
              <tbody>
                {activeAccounts.length === 0 ? (
                  <tr><td colSpan="5" className="p-6 text-center text-slate-500">No hay movimientos registrados para mostrar en el balance.</td></tr>
                ) : (
                  activeAccounts.map(acc => (
                    <tr key={acc.id} className="border-b border-slate-100 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800/50">
                      <td className="p-3 text-slate-800 dark:text-slate-200 border-r dark:border-slate-700">{acc.code} - {acc.name}</td>
                      <td className="p-3 text-right font-mono text-slate-800 dark:text-slate-200 border-r dark:border-slate-700">{formatCurrency(acc.debits, displayCurrency, exchangeRate)}</td>
                      <td className="p-3 text-right font-mono text-slate-800 dark:text-slate-200 border-r dark:border-slate-700">{formatCurrency(acc.credits, displayCurrency, exchangeRate)}</td>
                      <td className="p-3 text-right font-mono text-slate-800 dark:text-slate-200 border-r dark:border-slate-700">{acc.nature === 'Deudora' && acc.balance > 0 ? formatCurrency(acc.balance, displayCurrency, exchangeRate) : '-'}</td>
                      <td className="p-3 text-right font-mono text-slate-800 dark:text-slate-200">{acc.nature === 'Acreedora' && acc.balance > 0 ? formatCurrency(acc.balance, displayCurrency, exchangeRate) : '-'}</td>
                    </tr>
                  ))
                )}
                {activeAccounts.length > 0 && (
                  <tr className="bg-slate-50 dark:bg-slate-900 font-bold border-t-2 border-slate-200 dark:border-slate-600">
                    <td className="p-3 text-right text-slate-800 dark:text-slate-200 border-r dark:border-slate-700">TOTALES:</td>
                    <td className={`p-3 text-right font-mono border-r dark:border-slate-700 ${isSumBalanced ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-500'}`}>{formatCurrency(totalDebits, displayCurrency, exchangeRate)}</td>
                    <td className={`p-3 text-right font-mono border-r dark:border-slate-700 ${isSumBalanced ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-500'}`}>{formatCurrency(totalCredits, displayCurrency, exchangeRate)}</td>
                    <td className={`p-3 text-right font-mono border-r dark:border-slate-700 ${isBalBalanced ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-500'}`}>{formatCurrency(totalDebtorBalance, displayCurrency, exchangeRate)}</td>
                    <td className={`p-3 text-right font-mono ${isBalBalanced ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-500'}`}>{formatCurrency(totalCreditorBalance, displayCurrency, exchangeRate)}</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  const ViewIncomeStatement = () => (
    <div className="space-y-6 animate-in fade-in duration-300 max-w-4xl mx-auto">
      <h2 className="text-2xl font-bold text-slate-800 dark:text-white text-center">Estado de Resultados Integral</h2>

      <DateRangeFilter />

      <div className="flex justify-end">
        <button onClick={exportIncomeStatementToExcel} className="flex items-center gap-2 text-xs bg-emerald-100 hover:bg-emerald-200 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300 px-3.5 py-2 rounded-xl transition-colors font-medium">
          <FileSpreadsheet size={16} /> Exportar Excel
        </button>
      </div>

      <div className="bg-white dark:bg-slate-800 p-8 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 text-sm">
        <div className="space-y-4">
          <div className="flex justify-between items-end border-b border-slate-200 dark:border-slate-700 pb-2">
            <h3 className="font-bold text-slate-800 dark:text-white">Ingresos Operativos y Financieros</h3>
            <span className="font-mono font-bold text-slate-800 dark:text-white">{formatCurrency(financialStatements.revenue, displayCurrency, exchangeRate)}</span>
          </div>
          <div className="pl-4 space-y-2">
            {Object.values(accountBalances).filter(a => a.type === 'Ingreso' && a.category === 'Cuenta').map(acc => (
              <div key={acc.id} className="flex justify-between text-xs">
                <span className="text-slate-600 dark:text-slate-400">{acc.code} - {acc.name}</span>
                <span className="font-mono text-slate-700 dark:text-slate-300">{formatCurrency(acc.balance, displayCurrency, exchangeRate)}</span>
              </div>
            ))}
          </div>

          <div className="flex justify-between items-end border-b border-slate-200 dark:border-slate-700 pb-2 mt-6">
            <h3 className="font-bold text-slate-800 dark:text-white">Costos y Gastos</h3>
            <span className="font-mono font-bold text-slate-800 dark:text-white">({formatCurrency(financialStatements.expenses, displayCurrency, exchangeRate)})</span>
          </div>
          <div className="pl-4 space-y-2">
            {Object.values(accountBalances).filter(a => a.type === 'Gasto' && a.category === 'Cuenta').map(acc => (
              <div key={acc.id} className="flex justify-between text-xs">
                <span className="text-slate-600 dark:text-slate-400">{acc.code} - {acc.name}</span>
                <span className="font-mono text-slate-700 dark:text-slate-300">({formatCurrency(acc.balance, displayCurrency, exchangeRate)})</span>
              </div>
            ))}
          </div>

          <div className="flex justify-between items-center bg-slate-50 dark:bg-slate-900 p-4 rounded-xl mt-8 border border-slate-200 dark:border-slate-700">
            <h3 className="font-bold text-slate-800 dark:text-white">Utilidad (Pérdida) del Ejercicio</h3>
            <span className={`font-mono font-bold text-lg ${financialStatements.netIncome >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-500'}`}>
              {formatCurrency(financialStatements.netIncome, displayCurrency, exchangeRate)}
            </span>
          </div>
        </div>
      </div>
    </div>
  );

  const ViewBalanceSheet = () => (
    <div className="space-y-6 animate-in fade-in duration-300 max-w-5xl mx-auto">
      <h2 className="text-2xl font-bold text-slate-800 dark:text-white text-center">Estado de Situación Financiera</h2>

      <DateRangeFilter />
      <p className="text-xs text-slate-400 text-center -mt-3">Este reporte es una fotografía de saldos acumulados; se calcula usando los movimientos hasta la fecha "Hasta" indicada arriba.</p>

      <div className="flex justify-end">
        <button onClick={exportBalanceSheetToExcel} className="flex items-center gap-2 text-xs bg-emerald-100 hover:bg-emerald-200 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300 px-3.5 py-2 rounded-xl transition-colors font-medium">
          <FileSpreadsheet size={16} /> Exportar Excel
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8 text-sm">
        <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
          <h3 className="font-bold text-base text-slate-800 dark:text-white border-b-2 border-emerald-500 pb-2 mb-4">ACTIVOS</h3>
          <div className="space-y-3">
            {Object.values(accountBalances).filter(a => a.type === 'Activo' && a.category === 'Cuenta').map(acc => (
                <div key={acc.id} className="flex justify-between text-xs">
                  <span className="text-slate-600 dark:text-slate-400">{acc.code} - {acc.name}</span>
                  <span className="font-mono text-slate-800 dark:text-slate-200">{formatCurrency(acc.balance, displayCurrency, exchangeRate)}</span>
                </div>
            ))}
          </div>
          <div className="flex justify-between items-center mt-6 pt-4 border-t-2 border-slate-200 dark:border-slate-700">
            <h4 className="font-bold text-slate-800 dark:text-white">TOTAL ACTIVOS</h4>
            <span className="font-mono font-bold text-slate-800 dark:text-white">{formatCurrency(financialStatements.assets, displayCurrency, exchangeRate)}</span>
          </div>
        </div>

        <div className="space-y-8">
          <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
            <h3 className="font-bold text-base text-slate-800 dark:text-white border-b-2 border-red-500 pb-2 mb-4">PASIVOS</h3>
            <div className="space-y-3">
              {Object.values(accountBalances).filter(a => a.type === 'Pasivo' && a.category === 'Cuenta').map(acc => (
                  <div key={acc.id} className="flex justify-between text-xs">
                    <span className="text-slate-600 dark:text-slate-400">{acc.code} - {acc.name}</span>
                    <span className="font-mono text-slate-800 dark:text-slate-200">{formatCurrency(acc.balance, displayCurrency, exchangeRate)}</span>
                  </div>
              ))}
            </div>
            <div className="flex justify-between items-center mt-6 pt-4 border-t border-slate-200 dark:border-slate-700">
              <h4 className="font-bold text-slate-800 dark:text-white">Total Pasivos</h4>
              <span className="font-mono font-bold text-slate-800 dark:text-white">{formatCurrency(financialStatements.liabilities, displayCurrency, exchangeRate)}</span>
            </div>
          </div>

          <div className="bg-white dark:bg-slate-800 p-6 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700">
            <h3 className="font-bold text-base text-slate-800 dark:text-white border-b-2 border-blue-500 pb-2 mb-4">PATRIMONIO</h3>
            <div className="space-y-3">
              {Object.values(accountBalances).filter(a => a.type === 'Patrimonio' && a.category === 'Cuenta').map(acc => (
                  <div key={acc.id} className="flex justify-between text-xs">
                    <span className="text-slate-600 dark:text-slate-400">{acc.code} - {acc.name}</span>
                    <span className="font-mono text-slate-800 dark:text-slate-200">{formatCurrency(acc.balance, displayCurrency, exchangeRate)}</span>
                  </div>
              ))}
              <div className="flex justify-between text-xs italic">
                <span className="text-slate-600 dark:text-slate-400">Utilidad (Pérdida) del Ejercicio</span>
                <span className={`font-mono ${financialStatements.netIncome >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-red-500'}`}>
                  {formatCurrency(financialStatements.netIncome, displayCurrency, exchangeRate)}
                </span>
              </div>
            </div>
            <div className="flex justify-between items-center mt-6 pt-4 border-t border-slate-200 dark:border-slate-700">
              <h4 className="font-bold text-slate-800 dark:text-white">Total Patrimonio</h4>
              <span className="font-mono font-bold text-slate-800 dark:text-white">{formatCurrency(financialStatements.equity, displayCurrency, exchangeRate)}</span>
            </div>
          </div>

          <div className="bg-slate-900 text-white p-6 rounded-2xl shadow-sm">
            <div className="flex justify-between items-center">
              <h4 className="font-bold text-sm">TOTAL PASIVO + PATRIMONIO</h4>
              <span className="font-mono font-bold text-base">{formatCurrency(financialStatements.liabilities + financialStatements.equity, displayCurrency, exchangeRate)}</span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );

  const ViewInventory = () => {
    const totalInventoryValue = inventory.reduce((sum, item) => sum + item.totalValue, 0);

    return (
      <div className="space-y-6 animate-in fade-in duration-300">
        <div className="flex justify-between items-center">
          <h2 className="text-2xl font-bold text-slate-800 dark:text-white">Libro de Inventario</h2>
          <button onClick={exportInventoryToExcel} className="flex items-center gap-2 text-xs bg-emerald-100 hover:bg-emerald-200 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-300 px-3.5 py-2 rounded-xl transition-colors font-medium">
            <FileSpreadsheet size={16} /> Exportar Excel
          </button>
        </div>
        <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 overflow-hidden">
          <div className="p-4 border-b border-slate-100 dark:border-slate-700 flex justify-between items-center bg-slate-50 dark:bg-slate-900">
             <h3 className="font-bold text-sm text-slate-800 dark:text-white">Existencias y Repuestos</h3>
             <div className="text-right">
                <p className="text-xs text-slate-500">Valorización Total ({displayCurrency})</p>
                <p className="font-mono font-bold text-sm text-emerald-600 dark:text-emerald-400">{formatCurrency(totalInventoryValue, displayCurrency, exchangeRate)}</p>
             </div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-50 dark:bg-slate-900/50 border-b border-slate-200 dark:border-slate-700">
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300">Código Ref.</th>
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300">Artículo</th>
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300 text-right">Cantidad</th>
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300 text-right">Costo Unitario</th>
                  <th className="p-4 font-semibold text-slate-600 dark:text-slate-300 text-right">Costo Total</th>
                </tr>
              </thead>
              <tbody>
                {inventory.length === 0 ? (
                  <tr>
                    <td colSpan="5" className="p-6 text-center text-slate-500 italic">No hay artículos registrados en el inventario.</td>
                  </tr>
                ) : (
                  inventory.map(item => (
                    <tr key={item.id} className="border-b border-slate-100 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800/50">
                      <td className="p-4 text-slate-800 dark:text-slate-200 font-mono">{item.code}</td>
                      <td className="p-4 text-slate-800 dark:text-slate-200 font-medium">{item.name}</td>
                      <td className="p-4 text-slate-800 dark:text-slate-200 text-right">{item.quantity}</td>
                      <td className="p-4 text-slate-800 dark:text-slate-200 text-right font-mono">{formatCurrency(item.unitCost, displayCurrency, exchangeRate)}</td>
                      <td className="p-4 text-slate-800 dark:text-slate-200 text-right font-mono font-semibold">{formatCurrency(item.totalValue, displayCurrency, exchangeRate)}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  // Pantalla de carga mientras se leen los datos guardados.
  // (Va aquí, DESPUÉS de todos los useState/useMemo del componente, porque React exige
  // que los hooks se ejecuten siempre en el mismo orden en cada render.)
  if (!isDataLoaded) {
    return (
      <div className="flex h-screen items-center justify-center bg-slate-50 dark:bg-slate-950">
        <div className="flex flex-col items-center gap-3">
          <Loader2 className="w-8 h-8 text-emerald-600 animate-spin" />
          <p className="text-slate-500 text-sm">Conectando con la base de datos...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 font-sans transition-colors duration-200 overflow-hidden relative">
      {renderNav()}
      <main className="flex-1 overflow-y-auto p-6 md:p-8 lg:p-10 w-full">
        <div className="max-w-7xl mx-auto">
          {loadError && (
            <div className="mb-6 flex items-center gap-2 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 text-red-700 dark:text-red-300 text-sm px-4 py-3 rounded-xl">
              <AlertTriangle size={16} className="shrink-0" />
              {loadError}
            </div>
          )}
          {activeTab === 'dashboard' && <ViewDashboard />}
          {activeTab === 'cuentas' && <ViewChartOfAccounts />}
          {activeTab === 'diario' && <ViewJournal />}
          {activeTab === 'mayor' && <ViewLedger />}
          {activeTab === 'comprobacion' && <ViewTrialBalance />}
          {activeTab === 'resultados' && <ViewIncomeStatement />}
          {activeTab === 'situacion' && <ViewBalanceSheet />}
          {activeTab === 'inventario' && <ViewInventory />}
        </div>
      </main>
      
      <AccountModal />
      <ResetConfirmModal />
      <ExchangeRateModal />
    </div>
  );
}