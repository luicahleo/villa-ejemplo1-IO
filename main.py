# -*- coding: utf-8 -*-
"""
Created on Fri Nov  5 10:22:27 2021

@author: jgmir
"""

from ortools.linear_solver import pywraplp
# from IOfunctionsExcel import *

# name = 'EJEMPLO.xlsx'  # no ha subido el excel aunque parece muy sucio
# excel_doc = openpyxl.load_workbook(name, data_only=True)
# sheet = excel_doc['Hoja1']
#
# a = Read_Excel_to_List(sheet, 'B2', 'B5')
# b = Read_Excel_to_List(sheet, 'D2', 'D5')
# Fabricas = Read_Excel_to_List(sheet, 'A2', 'A5')
# Almacenes = Read_Excel_to_List(sheet, 'C2', 'C5')
# c = Read_Excel_to_NesteDic(sheet, 'F1', 'J5')


a=[8,7,6,2] #producción
b=[10,4,5,4] #demanda
c=[[3,1,4,5],[2,3,2,1],[1,4,5,3],[3,5,4,1]] #costes

#nos creamos nuestros indices, que vaya desde 1 hasta 4
Fabricas=[j for j in range(1,len(a)+1)]
Almacenes=[j for j in range(1,len(b)+1)]

def ejemplo():
    # declaramos el modelo
    solver = pywraplp.Solver.CreateSolver('GLOP')

    # declaramos variables, los cuales son diccionarios
    x = {}
    rfab = {}
    ralm = {}
    for i in Fabricas:
        x[i] = {}
        for j in Almacenes:
            x[i][j] = solver.NumVar(0, solver.infinity(), 'X%d;%d' % (i, j)) #ponemos limites, 0 e infinito, donde i es fabrica y j es almacen
    print('Número de variables=', solver.NumVariables()) #en este caso hay tantas variables como fabricas * demandas

    for i in Fabricas:
        rfab[i] = solver.Add(sum(x[i][j] for j in Almacenes) == a[i - 1], 'RF%d' % (i))

    for j in Almacenes:
        ralm[j] = solver.Add(sum(x[i][j] for i in Fabricas) == b[j - 1], 'RA%d' % (j))

    print('Número de restricciones=', solver.NumConstraints())

    # Ahora vamos con la funcion objetivo
    # donde 'c' es una lista de listas, en Minimize se hace suma(c*x)
    solver.Minimize(solver.Sum(c[i][j]*x[i][j] for i in Fabricas for j in Almacenes))  # c[i-1][j-1] para listas internas

    # recogemos el resultado del solver
    # status = solver.Solve()
    #
    # if status == pywraplp.Solver.OPTIMAL:
    #     for i in Fabricas:
    #         for j in Almacenes:
    #             print('X%d;%d = %d' % (i, j, x[i][j].solution_value()))
    #     # for i in Fabricas:
    #     #     for j in Almacenes:
    #     #         print('CR%d;%d = %d' % (i, j, x[i][j].ReducedCost()))
    #     # for i in Fabricas:
    #     #     print('u%d=%d' % (i, rfab[i].dual_value()))
    #     # for j in Almacenes:
    #     #     print('v%d=%d' % (j, ralm[j].dual_value()))
    #     print('Funcion objetivo =', solver.Objective().Value())
    #
    # else:
    #     print('El problema es inadmisible')

    # Solu = {}
    # for i in Almacenes:
    #     Solu[i] = {j: 0.0 for j in Almacenes}
    # for i in Fabricas:
    #     for j in Almacenes:
    #         Solu[i][j] = x[i][j].solution_value()
    # Write_NesteDic_to_Excel(excel_doc, name, sheet, Solu, 'F8', 'J12')


ejemplo()