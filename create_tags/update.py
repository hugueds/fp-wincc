prefix = 'UPT_BIO_POSTO[i]_'

sufix = [
        'IDENTIFIED\tBinary Tag\tDB205,DX.0',
        'LEVEL\tUnsigned 8-bit value\tDB205,DBBX',
        'NEW_LEVEL\tUnsigned 8-bit value\tDB205,DBBX',
        'REG\tUnsigned 32-bit value\tDB205,DDX",',
        'SSB\tText tag 8-bit character set\tDB205,DBBX'
        ]

size = 16
offset = [0,1,2,4,8] #16, 17, 18, 20, 24 -- SIZE = 16

stations = 27

for i in range(stations+1):
    k = 0
    for j in sufix:
        print prefix.replace('i', str(i))+j.replace('X', str(offset[k]))
        k += 1
    offset = map(lambda x: x + size, offset)
