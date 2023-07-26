
"""
Author: Yu-Sheng Tang 4765946

Uni Freiburg

Sustainable Systems Engineering

Winter Semester 19, 20

Operations Research Project

Using the data from the paper and implementing with AHP Method

"""

import numpy as np
import xlrd

# =============================================================================

# Read the data from the xlsx

Sheet_dict = {
0: 'Pairwise Comparison',
1: 'LCOE',
2: 'Ability to respond to demand',
3: 'Efficiency',
4: 'Capacity factor', 
5: 'Land use',
6: 'External costs(Environmental)',
7: 'External costs(Human health)',
8: 'Job creation',
9: 'Social acceptability', 
10: 'External supply risk'
}
#use Sheet_dict[2] to call the value in the dictionary

Matrices=[]
for i in range(0,11):
    book = xlrd.open_workbook(r'Tang_Roland_AHP_InputMatrices.xlsx')
    sheet = book.sheet_by_name(Sheet_dict[i])
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

    Matrices.append(data) 

#    print()
#    print(str(Sheet_dict[i]))
#    print()
#    print(data)
#    print()

# =============================================================================


RI_dict = {1: 0, 2: 0, 3: 0.5245, 4: 0.8815, 5: 1.1086,
           6: 1.2479, 7: 1.3417, 8: 1.4056, 9: 1.4499, 10:1.4857,
           11:1.5141, 12:1.5365, 13:1.5551, 14:1.5713, 15:1.5838}
#Random Index form Alonso, Lamata (2006)

def get_w(array):
    row = array.shape[0]  # Calculate the order of the matrix
    a_axis_0_sum = array.sum(axis=0)
    # print(a_axis_0_sum)
    b = array / a_axis_0_sum  # New matrix b
    # print(b)
#    b_axis_0_sum = b.sum(axis=0)
    b_axis_1_sum = b.sum(axis=1)  # eigenvector for each column
    # print(b_axis_1_sum)
    w = b_axis_1_sum / row  # Normalization(eigenvector)
#    nw = w * row
    AW = (w * array).sum(axis=1)
    # print(AW)
    max_max = sum(AW / (row * w))
    # print(max_max)
    CI = (max_max - row) / (row - 1)
    CR = CI / RI_dict[row]
    if CR < 0.1:
        # print(round(CR, 3))
        # print('Sufficient Normalization')
        # print(np.max(w))
        # print(sorted(w,reverse=True))
        # print(max_max)
        # print('Eigenvector:%s' % w)
#        print(round(CR, 3))
        return w
        
    else:
        print(round(CR, 3))
        print('Insufficient Normalization, please modified the matrix!')


def main(array):
    if type(array) is np.ndarray:
        return get_w(array)
    else:
        print('Please input an array!')


if __name__ == '__main__':
   
    #all the matrices are calculate in the excel file 
    
    #Pairwise Comparison Matrix 
    e = np.array(Matrices[0])
    #LCOE
    a = np.array(Matrices[1])
    #Ability to respond to demand
    b = np.array(Matrices[2])
    #Efficiency
    c = np.array(Matrices[3])
    #Capacity factor 
    d = np.array(Matrices[4])   
    #Land use
    f = np.array(Matrices[5])   
    #External costs(Environmental)
    g = np.array(Matrices[6])    
    #External costs(Human health)
    h = np.array(Matrices[7])    
    #Job creation
    i = np.array(Matrices[8])    
    #Social acceptability
    j = np.array(Matrices[9])   
    #External supply risk
    k = np.array(Matrices[10])   
    

    e = main(e) #pairwise comparison matrix
    
    a = main(a) #score matrix a : LCOE
    b = main(b) #score matrix b : Ability to respond to demand
    c = main(c) #score matrix c : Efficiency
    d = main(d) #score matrix d : Capacity factor
    
    f = main(f) #score matrix f : Land use
    g = main(g) #score matrix g : External costs(Environmental)
    h = main(h) #score matrix h : External costs(Human health)
    i = main(i) #score matrix i : Job creation
    j = main(j) #score matrix j : Social acceptability
    k = main(k) #score matrix k : External supply risk
    
    try:
        res = np.array([a, b, c, d, f, g, h, i, j, k])
        ret = (np.transpose(res) * e).sum(axis=1)
        print(ret)
        
    except TypeError:
        print('The matrix is wrong, e.g. insufficient Normalization, please modify the data!')