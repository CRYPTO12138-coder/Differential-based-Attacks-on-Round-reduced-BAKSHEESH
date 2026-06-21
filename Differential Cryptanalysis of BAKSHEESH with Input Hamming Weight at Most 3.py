from gurobipy import *

def GetVariables(round,varName,varSize,variable):
    res = []
    for i in range(varSize):
        res.append(varName + "_" + str(round) + "_" + str(i))
        variable.add(varName + "_" + str(round) + "_" + str(i))
    return res

##################################################differential##################################################
def Constraint_initialize_D(R_D, f, variable):
    f.write("c" + " = 1 " + "\n")

    res = []
    for i in range (STATE_LENGTH):
        res.append(GetVariables(0,"X",STATE_LENGTH,variable)[i])
    f.write(" + ".join(res) + " >= 1 " + "\n")

    res = []
    for i in range (STATE_LENGTH):
        res.append(GetVariables(0,"X",STATE_LENGTH,variable)[i])
    f.write(" + ".join(res) + " <= 3 " + "\n")

def Constraint_sbox_D(r, f, variable):
    P = [0, 33, 66, 99, 96, 1, 34, 67, 64, 97, 2, 35, 32, 65, 98, 3, 4, 37, 70, 103, 100, 5, 38, 71, 68, 101, 6, 39, 36, 69, 102, 7, 8, 41, 
    74, 107, 104, 9, 42, 75, 72, 105, 10, 43, 40, 73, 106, 11, 12, 45, 78, 111, 108, 13, 46, 79, 76, 109, 14, 47, 44, 77, 110, 15, 16, 49, 
    82, 115, 112, 17, 50, 83, 80, 113, 18, 51, 48, 81, 114, 19, 20, 53, 86, 119, 116, 21, 54, 87, 84, 117, 22, 55, 52, 85, 118, 23, 24, 57, 
    90, 123, 120, 25, 58, 91, 88, 121, 26, 59, 56, 89, 122, 27, 28, 61, 94, 127, 124, 29, 62, 95, 92, 125, 30, 63, 60, 93, 126, 31]

    M0 = [[2, 1, -3, -3, -2, 2, -1, -1, 7, 0], [-3, -2, 0, -2, 2, 1, 1, -1, 6, 0], [-1, 0, 0, 0, -2, -1, -1, -1, -2, -6], [1, 0, 0, 0, 2, 1, 1, 1, -2, 0], [-1, 2, 0, 0, 4, -1, -1, -1, 2, 0], [1, 2, 0, 0, -4, 1, 1, 1, 2, 0], [0, 1, 1, 1, 0, 0, 0, 0, -1, 0], [1, 1, 0, 3, 2, -2, -2, 1, 1, 0], [2, 1, 3, 0, -2, -1, 2, -1, 1, 0], [-1, 1, 0, 3, -2, 2, 2, -1, 1, 0], [-2, 1, 3, 0, 2, 1, -2, 1, 1, 0], [-1, -2, -2, 0, -2, 1, -1, 1, 4, -2], [1, -2, 0, -2, -2, -1, -1, 1, 4, -2], [1, -2, -2, 0, 2, -3, 1, -1, 6, 0], [-2, 1, -3, -3, 2, -2, 1, 1, 7, 0], [0, -1, -1, 1, 1, 1, -1, -1, 3, 0], [1, -1, 1, -1, 1, 0, -1, -1, 3, 0], [-1, 0, 1, 1, 1, -1, 1, 0, 1, 0], [1, 0, 1, 1, -1, 1, -1, 0, 1, 0], [-1, -1, 1, -1, -1, 0, 1, 1, 3, 0], [0, -1, -1, 1, -1, -1, 1, 1, 3, 0]]

    for i in range (int(STATE_LENGTH/4)):
        for t in range (len(M0)):
            res = []
            res.append(str(M0[t][0]) + " " + GetVariables(r,"X",STATE_LENGTH,variable)[4*i+3])
            res.append(str(M0[t][1]) + " " + GetVariables(r,"X",STATE_LENGTH,variable)[4*i+2])
            res.append(str(M0[t][2]) + " " + GetVariables(r,"X",STATE_LENGTH,variable)[4*i+1])
            res.append(str(M0[t][3]) + " " + GetVariables(r,"X",STATE_LENGTH,variable)[4*i+0])
            res.append(str(M0[t][4]) + " " + GetVariables(r+1,"X",STATE_LENGTH,variable)[P[4*i+3]])
            res.append(str(M0[t][5]) + " " + GetVariables(r+1,"X",STATE_LENGTH,variable)[P[4*i+2]])
            res.append(str(M0[t][6]) + " " + GetVariables(r+1,"X",STATE_LENGTH,variable)[P[4*i+1]])
            res.append(str(M0[t][7]) + " " + GetVariables(r+1,"X",STATE_LENGTH,variable)[P[4*i+0]])
            res.append(str(M0[t][8]) + " " + GetVariables(r,"p_D",int(STATE_LENGTH/4),variable)[i])
            f.write(" + ".join(res) + " - " + str(M0[t][9]) + " c" + " >= 0 " + "\n")

def Constraint_D(f, variable):
    Constraint_initialize_D(R_D, f, variable)
    for r in range (0, R_D):
        Constraint_sbox_D(r, f, variable)

def ObjectiveFunction_D(f, variable):
    res = []
    for r in range (0, R_D):
        for i in range (STATE_LENGTH):
            res.append("2 " + GetVariables(r,"p_D",STATE_LENGTH,variable)[i])
    f.write(" + ".join(res) + "\n")

def VariablesType(f):
    f.write("\n".join(variable) + "\n")

def CreateModel(lpFileName, variable):
    f = open(lpFileName, "w")
    f.write("Minimum\n")
    ObjectiveFunction_D(f, variable)
    f.write("Subject To\n")
    Constraint_D(f, variable)
    f.write("Binaries\n")
    VariablesType(f)
    f.write("End\n")
    f.close()

def SolveModel(lpFileName, solFileName):
    model = read(lpFileName)
    # model.setParam('OutputFlag', 0)
    # model.Params.PoolSearchMode = 2
    # model.Params.PoolSolutions = 200000000
    model.optimize()
    # model.computeIIS()
    # model.write("LELBC_dl.ilp")
    model.write(solFileName)
    # if model.status == GRB.OPTIMAL:
    #     print('Optimal objective:', model.objVal)
    #     print(model.status)


if __name__ == '__main__':

    STATE_LENGTH = 128

    R_D = 7

    variable = set()
    lpFileName = "BAKSHEESH_d.lp"
    solFileName = "BAKSHEESH_d_%d.sol" % R_D
    CreateModel(lpFileName, variable)
    SolveModel(lpFileName, solFileName)
