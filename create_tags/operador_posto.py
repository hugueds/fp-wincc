prefix = "DB_OPERADOR_POSTO[s]"
sufixStation = [
            "_CMD_W_AccessLevel\tSigned 16-bit value\tDB175,DBWX",
            "_CMD_W_Station\tSigned 16-bit value\tDB175,DBWX",
            "_CMD_W_Position\tSigned 16-bit value\tDB175,DBWX",
            "_CMD_WR_Trigger\tSigned 16-bit value\tDB175,DBWX",
            "_CMD_R_timeCounter\tSigned 16-bit value\tDB175,DBWX",     
            "_CMD_R_invalidCounter\tSigned 16-bit value\tDB175,DBWX",
        ]

sufixPosition = [
            "POSICAO[p]_IDENTIFICADO\tBinary Tag\tDB175,DX.0",
            "POSICAO[p]_HABILITADO\tBinary Tag\tDB175,DX.1",
            "POSICAO[p]_ID_WORKSTATION\tUnsigned 8-bit value\tDB175,DBBX",
            "POSICAO[p]_TRAINING_LEVEL\tSigned 16-bit value\tDB175,DBWX",
            "POSICAO[p]_SSB\tText tag 8-bit character set\tDB175,DBBX",
            "POSICAO[p]_REGISTRO\tUnsigned 32-bit value\tDB175,DDX"
        ]
         

offsetStation = [0,2,4,6,8,10]
offsetPosition = [12,12,13,14,16,24]

stations = 27

for i in range(stations+1):
    k = 0
    l = 0
    pos = 1
    for j in sufixStation:
        print prefix.replace('s', str(i))+j.replace('X', str(offsetStation[k]))
        k += 1        
    for l in range(6):
        m = 0
        for pos in sufixPosition:
            print prefix.replace('s', str(i)) + pos.replace('p', str(l+1)).replace('X', str(offsetPosition[m]))
            m += 1
        offsetPosition = map(lambda x: x + 16, offsetPosition)
    offsetPosition = map(lambda x: x + 12, offsetPosition)
    offsetStation = map(lambda x: x + 108, offsetStation)
        
    
