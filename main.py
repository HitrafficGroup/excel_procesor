from cal030 import calcular030
from cal020 import calcular020
from cal040 import calcular040
from cal050A import calcular050A
from cal050M import calcular050M
from calcular_porcentaje import calcularPorcent
while True:
    print("[---Menu---]")
    print("1) calcular BD 020")
    print("2) calcular BD 030")
    print("3) calcular BD 040")
    print("4) calcular BD 050A")
    print("5) calcular BD 050M")
    print("6) calcular porcentaje% ")
    #print("7) calcular Requerimiento ")
    print("9) Salir")
    opcion = input("Ingrese la opcion: ")
    if opcion == "1":
        print("calculando 020 ........")
        calcular020()
    elif opcion == "2":
        print("calculando 030 ........")
        calcular030()
    elif opcion == "3":
        print("calculando 040 ........")
        calcular040()
    elif opcion == "4":
        print("calculando 050A ........")
        calcular050A()
    elif opcion == "5":
        print("calculando 050M ........")
        calcular050M()
    elif opcion == "6":
        print("[---Escoga la opcion de cual quiere determinar el porcentaje ---]")
        print("1)  BD 020")
        print("2)  BD 030")
        print("3)  BD 040")
        print("4)  BD 050A")
        print("5)  BD 050M")
        print("6)  cualquier numero  para cancelar")
        x = input("Ingrese la opcion: ")
        if x == "1": 
            #porcentaje = input("Ingrese el porcentaje en formato decimal ejemplo 0.02 : ")
            calcularPorcent("resultados/cal_020_bd.xlsx",'0.08','_020')
            print("calculando porcentaje % ........")    
        elif x == "2":
            #porcentaje = input("Ingrese el porcentaje en formato decimal ejemplo 0.02 : ")
            calcularPorcent("resultados/cal_030_bd.xlsx",'0.08','_030')
            print("calculando porcentaje % ........")
        elif x == "3":
            #porcentaje = input("Ingrese el porcentaje en formato decimal ejemplo 0.02 : ")
            calcularPorcent("resultados/cal_040_bd.xlsx",'0.08','_040')
            print("calculando porcentaje % ........")
        elif x == "4":
            #porcentaje = input("Ingrese el porcentaje en formato decimal ejemplo 0.02 : ")
            calcularPorcent("resultados/cal_050M_bd.xlsx",'0.08','_050A')
            print("calculando porcentaje % ........")
        elif x == "5":
            #porcentaje = input("Ingrese el porcentaje en formato decimal ejemplo 0.02 : ")
            calcularPorcent("resultados/cal_051A_bd.xlsx",'0.08','_050M')
            print("calculando porcentaje % ........")
        else:
            pass
    elif opcion == "7":
        print("calculando Requerimiento   ........")
    elif opcion == "9":
        break