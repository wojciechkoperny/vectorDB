import cantools
import xlsxwriter
import os.path
import sys

db = []

def parseVectorDB(path):
    global db
    try:
        if os.path.exists(path):
            db = cantools.database.load_file(path)
            print('Parsing file: '+ path)
        else:
            raise ValueError
    except ValueError:
        print('dbc under path not found - please provide proper proper dbc file as imput parameter')
        sys.exit(1)

def prepareWorkbook():

    try:
        workbook = xlsxwriter.Workbook('C2_Body.xlsx')
    except xlsxwriter.exceptions.XlsxFileError:
        print('please close output workbook before running a script')
        sys.exit(1)

    header_format = workbook.add_format({'bold': True,
                                     'align': 'center',
                                     'valign': 'vcenter',
                                     'fg_color': '#D7E4BC',
                                     'border': 1})
 
    worksheetMessages = workbook.add_worksheet("Messages")
    worksheetNodes = workbook.add_worksheet("Nodes")

    row = 0
    col = 0

    worksheetNodes.write(row,col,"NODE")
    worksheetNodes.write(row,col+1,"NODE COMMENT")

    row+=1

    for node in db.nodes:
        worksheetNodes.write(row, col, node.name)
        worksheetNodes.write(row, col+1, node.comment)
        row+=1

    row = 0
    col = 0

    worksheetMessages.write(row, col ,"MESSAGE",header_format)
    worksheetMessages.write(row,col+1,"MESSAGE ID",header_format)
    worksheetMessages.write(row,col+2,"SENER",header_format)
    worksheetMessages.write(row,col+3,"SEND TYPE",header_format)
    worksheetMessages.write(row,col+4,"CYCLE",header_format)
    worksheetMessages.write(row,col+5,"MESSAGE LENGTH",header_format)
    worksheetMessages.write(row,col+6,"SIGNAL NAME",header_format)
    worksheetMessages.write(row,col+7,"SIGNAL START BIT",header_format)
    worksheetMessages.write(row,col+8,"SIGNAL LENGTH",header_format)
    worksheetMessages.write(row,col+9,"SIGNAL MAX",header_format)
    worksheetMessages.write(row,col+10,"SIGNAL MIN",header_format)
    worksheetMessages.write(row,col+11,"SIGNAL OFFSER",header_format)
    worksheetMessages.write(row,col+12,"SIGNAL SCALE",header_format)
    worksheetMessages.write(row,col+13,"SIGNAL UNIT",header_format)
    worksheetMessages.write(row,col+14,"SIGNAL RECEIVER",header_format)
    row=1

    worksheetMessages.freeze_panes(1, 0)

    for message in db.messages:
        for signal in message.signals:
            worksheetMessages.write(row,col,message.name)
            worksheetMessages.write_formula(row,col+1,'="0x"&DEC2HEX('+str(message.frame_id)+')')
            worksheetMessages.write(row,col+2,message.senders[0])
            worksheetMessages.write(row,col+3,message.send_type)
            worksheetMessages.write(row,col+4,message.cycle_time)
            worksheetMessages.write(row,col+5,message.length)
            worksheetMessages.write(row,col+6,signal.name)
            worksheetMessages.write(row,col+7,signal.start)
            worksheetMessages.write(row,col+8,signal.length)
            worksheetMessages.write(row,col+9,signal.maximum)
            worksheetMessages.write(row,col+10,signal.minimum)
            worksheetMessages.write(row,col+11,signal.offset)
            worksheetMessages.write(row,col+12,signal.scale)
            worksheetMessages.write(row,col+13,signal.unit)
            worksheetMessages.write(row,col+14,str(signal.receivers))
            row+=1
    try:
        workbook.close()
    except xlsxwriter.exceptions.XlsxFileError:
        print('please close output workbook before running a script')
        sys.exit(1)
    


if __name__ == '__main__':

    try:
        parseVectorDB(sys.argv[1])
        prepareWorkbook()
        print('Communication matrix was succesfully generated!')
    
    except IndexError:
        print('Please provide path to dbc file as imput argument')
        sys.exit(1)
    except Exception as e:
        print('General error: ' + str(e) + '.\nShutting down.')
        sys.exit(1)
    
