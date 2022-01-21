import cantools
from pprint import pprint
import xlsxwriter

db

def parseVectorDB():

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
worksheetMessages.write(row,col+6,"SIGNAL NAME")
worksheetMessages.write(row,col+7,"SIGNAL START BIT")
worksheetMessages.write(row,col+8,"SIGNAL LENGTH")
worksheetMessages.write(row,col+9,"SIGNAL MAX")
worksheetMessages.write(row,col+10,"SIGNAL MIN")
worksheetMessages.write(row,col+11,"SIGNAL OFFSER")
worksheetMessages.write(row,col+12,"SIGNAL SCALE")
worksheetMessages.write(row,col+13,"SIGNAL UNIT")
worksheetMessages.write(row,col+14,"SIGNAL RECEIVER")



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
        #worksheetMessages.write_formula((row,col,'='+str(message.length)))
        worksheetMessages.write(row,col,message.length)
        col+=1
        worksheetMessages.write(row,col,signal.name)
        col+=1
        worksheetMessages.write(row,col,signal.start)
        col+=1
        worksheetMessages.write(row,col,signal.length)
        col+=1
        worksheetMessages.write(row,col,signal.maximum)
        col+=1
        worksheetMessages.write(row,col,signal.minimum)
        col+=1
        worksheetMessages.write(row,col,signal.offset)
        col+=1
        worksheetMessages.write(row,col,signal.scale)
        col+=1
        worksheetMessages.write(row,col,signal.unit)
        col+=1
        worksheetMessages.write(row,col,str(signal.receivers))
        col+=1

        col=0
        row+=1
   #worksheetMessages.write()

print("proces")
example_message = db.get_message_by_name('SWCU_REQ')
#pprint(example_message.signals)
print("finish")

workbook.close()


if __name__ == '__main__':
    try:
        parseVectorDB()
    except:
        print("cos poszlo nie tak")
