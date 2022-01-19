import cantools
from pprint import pprint
import xlsxwriter


db = cantools.database.load_file('C2_Body.dbc')
workbook = xlsxwriter.Workbook('C2_Body.xlsx')
#print(type(db.buses))
#for bus in range(len(db.buses)):
worksheetNodes = workbook.add_worksheet("Nodes")
worksheetMessages = workbook.add_worksheet("Messages")




#print(type(worksheet))
print(type(db.messages))

pprint(db.nodes)

row = 0
col = 0

worksheetNodes.write(row,col,"NODE")
worksheetNodes.write(row,col+1,"NODE COMMENT")

row+=1

for node in range(len(db.nodes)):
    worksheetNodes.write(row+node, col, db.nodes[node].name)
    worksheetNodes.write(row+node, col+1, db.nodes[node].comment)
    node += 1

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
for message in range(len(db.messages)):
    worksheetMessages.write(row+message,col,db.messages[message].name)
    worksheetMessages.write(row+message,col+1,db.messages[message].frame_id)
    worksheetMessages.write(row+message,col+2,db.messages[message].senders[0])
    worksheetMessages.write(row+message,col+3,db.messages[message].send_type)
    worksheetMessages.write(row+message,col+4,db.messages[message].cycle_time)
    worksheetMessages.write(row+message,col+5,db.messages[message].length)


print("proces")
example_message = db.get_message_by_name('SWCU_REQ')
#pprint(example_message.signals)
print("finish")

workbook.close()