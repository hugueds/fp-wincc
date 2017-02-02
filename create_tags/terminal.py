prefix = 'TERM[i]'

sufix = ['_ID_WORKSTATION\tUnsigned 8-bit value\tDB119,DBBX',
         '_IDENTIFICADO\tBinary Tag\tDB119,DX.0',         
         '_HABILITADO\tBinary Tag\tDB119,DX.1',
         '_SSB\tText tag 8-bit character set\tDB119,DBBX',
         '_REGISTRO\tUnsigned 32-bit value\tDB119,DDX',    
         '_NIVEL_ACESSO\tUnsigned 16-bit value\tDB119,DBWX'         
         ]

offset = [0,1,1,2,10,14] #16, 17, 17, 18, 26, 30

stations = 27

for i in range(stations+1):
    k = 0
    for j in sufix:
        print prefix.replace('i', str(i))+j.replace('X', str(offset[k]))
        k += 1
    offset = map(lambda x: x + 16, offset)
