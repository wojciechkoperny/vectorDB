import cantools
from pprint import pprint
import xlsxwriter


db = cantools.database.load_file('C2_Body.dbc')
workbook = xlsxwriter.Workbook('C2_Body.xlsx')
#print(type(db.buses))
#for bus in range(len(db.buses)):
worksheetMessages = workbook.add_worksheet("Messages")
worksheetNodes = workbook.add_worksheet("Nodes")




#print(type(worksheet))
#print(type(db.messages))

#pprint(db.nodes)

row = 0
col = 0

worksheetNodes.write(row,col,"NODE")
worksheetNodes.write(row,col+1,"NODE COMMENT")

row+=1

for node in db.nodes:
    worksheetNodes.write(row, col, node.name)
    worksheetNodes.write(row, col+1, node.comment)
    row+=1
    #node += 1

#    for

row = 0
col = 0

worksheetMessages.write(row, col ,"MESSAGE")
worksheetMessages.write(row,col+1,"MESSAGE ID")
worksheetMessages.write(row,col+2,"SENER")
worksheetMessages.write(row,col+3,"SEND TYPE")
worksheetMessages.write(row,col+4,"CYCLE")
worksheetMessages.write(row,col+5,"MESSAGE LENGTH")

row=1
for message in db.messages:
    for signal in message.signals:
        
        worksheetMessages.write(row,col,message.name)
        col+=1
        worksheetMessages.write_formula(row,col,'="0x"&DEC2HEX('+str(message.frame_id)+')')
        col+=1
        worksheetMessages.write(row,col,message.senders[0])
        col+=1
        worksheetMessages.write(row,col,message.send_type)
        col+=1
        worksheetMessages.write(row,col,message.cycle_time)
        col+=1
        worksheetMessages.write(row,col,message.length)
        col+=1
        worksheetMessages.write(row,col,signal.name)
        col=0
        row+=1
   #worksheetMessages.write()

print("proces")
example_message = db.get_message_by_name('SWCU_REQ')
#pprint(example_message.signals)
print("finish")

workbook.close()