# Author - Nitin Bhagat

# this file should read all the inputs needed for server mapping info

# Database name
# Number of instances
# I/O Growth - 10% by default
# Server - hostname
# Read requests/sec - max of read_iops_max - if there are multiple instances, read ops value/ no of instaces
# Write requests/sec - max of write_iops_max - if there are multiple instances, write ops value/ no of instaces
# DB Avg CPU Util
# SGA Size GB - Max value of SGA column under Begin Memory
# PGA Size GB - Max value of PGA column under Begin Memory
import os
import sys
import re
import csv
import math
import openpyxl
import win32com.client
# import ReadMacros

stage_directory = sys.argv[1]
outFilesString = sys.argv[2]
environmentName = sys.argv[3]
IOGrowth = sys.argv[4]
workbookFile = sys.argv[5]
outFiles = list(outFilesString.split(" "))

row = [[]]


# Round function for consistent rounding
def normal_round(n):
    if n - math.floor(n) < 0.5:
        return math.floor(n)
    return math.ceil(n)

# Function to assign 0 to a variable if empty
def assignZero(n):
    if len(n)==0:
        return 0
    else:
        return n


def insertToCSV(databaseName, instances, server, IOGrowth, read_iops, write_iops, sga, pga):
    row.append([databaseName, instances, server, IOGrowth, read_iops, write_iops, sga, pga])


def getAllValues(f1):
    
    # all_sga = {}
    # all_pga = {}
    print(".")
    for line in f1:
        if re.search(r'^DB_NAME', line):
            databaseName = line
            databaseName = databaseName[8:]
            databaseName = databaseName.strip()
        if re.search(r'^INSTANCES', line):
            instances = line
            instances = instances[9:]
            instances = int(instances.strip())
        if re.search(r'^HOSTS', line):
            server = line
            server = server[5:]
            server = server.strip()
            server_list = server.split(',')
    print("Database Name: %s" % databaseName)
    print("Instances: %d" % instances)
    print("Host: %s" % server)
    # for servers in server_list:
    #     # if servers not in all_read_iops:
    #     #     # all_read_iops[servers] = []
    #     #     # all_read_iops.append(servers)
    #     # if servers not in all_write_iops:
    #     #     # all_write_iops[servers] = []
    #     #     # all_write_iops.append(servers)
    #     if servers not in all_sga:
    #         all_sga[servers] = []
    #     if servers not in all_pga:
    #         all_pga[servers] = []

    #####################################################################################################
    # get BEGIN-MAIN-METRICS & END-MAIN-METRICS line number
    begin_main_metrics = 0
    end_main_metrics = 0
    iops_check_string = 0
    for num, line in enumerate(f1, 1):
        if '~~BEGIN-MAIN-METRICS~~' in line:
            begin_main_metrics = num + 3
            iops_check_string = num + 1 # Line where string exists
        if '~~END-MAIN-METRICS~~' in line:
            end_main_metrics = num-2
            break
    start_read_iops_max = f1[iops_check_string].find("read_iops") + 1 # Find the starting position of string "read_iops_max"
    end_read_iops_max = start_read_iops_max + len("read_iops") - 1 # Find the ending position of string "read_iops_max"
    
    start_write_iops_max = f1[iops_check_string].find("write_iops") + 1 # Find the starting position of string "write_iops_max"
    end_write_iops_max = start_write_iops_max + len("write_iops") - 1 # Find the ending position of string "write_iops_max"
    
    # iops_start_inst = f1[iops_check_string].find("inst") + 1 # Find the starting position of string "inst"
    # iops_end_inst = iops_start_inst + 9 # Find the ending position of string "inst"

    start_snap = f1[iops_check_string].find("snap") - 4 # Find the starting position of string "snap"
    end_snap = start_snap + 9 # Find the ending position of string "snap"

    read_iops_list = {}
    write_iops_list = {}
    # all_read_iops = []
    # all_write_iops = []
    max_read_iops = 0
    max_write_iops = 0
    # Iterate through the above 2 lines
    for i in range(begin_main_metrics, end_main_metrics):
        # Gets the entire line
        line = f1[i]
        # Slices the line to include only read_iops and write_iops. Replace ',' with '.' (if file is corrupted)
        read_iops = line[start_read_iops_max:end_read_iops_max]
        read_iops = read_iops.replace(',', '.')
        write_iops = line[start_write_iops_max:end_write_iops_max]
        write_iops = write_iops.replace(',', '.')
        # Removes extra spaces 
        read_iops = read_iops.strip()
        write_iops = write_iops.strip()
        # Convert from string to float
        read_iops = float(read_iops)
        write_iops = float(write_iops)
        # Do above for snap, add to dictionary as key and give a initial value of 0
        snap = line[start_snap:end_snap]
        snap = snap.replace(',', '.')
        snap = snap.strip()
        snap = int(snap)
        if snap not in read_iops_list:
            read_iops_list[snap] = 0
        if snap not in write_iops_list:
            write_iops_list[snap] = 0
        # read_iops_list.append(read_iops)
        # write_iops_list.append(write_iops)

        # Add read_iops and write_iops to list for the particular snap ID
        read_iops_list[snap] += read_iops
        write_iops_list[snap] += write_iops

        # Slices the line to include only instance
        # instance_number = line[iops_start_inst:iops_end_inst]
        # Removes extra spaces
        # instance_number = instance_number.strip()
        # Convert from string to float
        # instance_number = int(instance_number)

        # Append the read_iops_max to the corresponding key (ServerName)
        # all_read_iops[server_list[instance_number-1]].append(normal_round(read_iops/instances))
        # all_read_iops.append(read_iops)

        # Append the write_iops_max to the corresponding key (ServerName)
        # all_write_iops[server_list[instance_number-1]].append(normal_round(write_iops/instances))
        # all_write_iops.append(write_iops)

    # Get maximum of read_iops and write_iops grouped by snap
    keymax_read_iops = max(read_iops_list.keys(), key=(lambda k: read_iops_list[k]))
    max_read_iops = read_iops_list[keymax_read_iops]
    keymax_write_iops = max(write_iops_list.keys(), key=(lambda k: write_iops_list[k]))
    max_write_iops = write_iops_list[keymax_write_iops]

    # print("ALL READ IOPS")
    # print(read_iops_list)

    # print("ALL WRITE IOPS")
    # print(write_iops_list)

    # print("MAX READ IOPS")
    # print(max_read_iops)

    # print("MAX WRITE IOPS")
    # print(max_write_iops)

    # Iterate through list to sort according to number of instances
    # max_read_iops = []
    # max_write_iops = []

    # Read every i'th object in list
    # for i in range(instances):
    #     max_read_iops.append(normal_round(max(read_iops_list[::(i+1)])))
    #     max_write_iops.append(normal_round(max(write_iops_list[::(i+1)])))

    # # Divide list by the number of instances
    # max_read_iops = [normal_round(i/instances) for i in max_read_iops]
    # max_write_iops = [normal_round(i/instances) for i in max_write_iops]

    # # Print max_read_iops & max_write_iops for every instance
    # for i in range(instances):
    #     print("read_iops_max for instance %s: %d" %
    #           (server_list[i], max_read_iops[i]))
    #     print("write_iops_max for instance %s: %d" %
    #           (server_list[i], max_write_iops[i]))

    #####################################################################################################
    # get BEGIN-MEMORY and END-MEMORY line number
    begin_memory = 0
    end_memory = 0
    check_string = 0
    for num, line in enumerate(f1, 1):
        if '~~BEGIN-MEMORY~~' in line:
            begin_memory = num + 3
            check_string = num + 1 # Line where string exists
        if '~~END-MEMORY~~' in line:
            end_memory = num-2
            break
    
    start_sga = f1[check_string].find("SGA") - 3 # Find the starting position of string "SGA"
    end_sga = start_sga + len("SGA") + 3 # Find the ending position of string "SGA"

    start_pga = f1[check_string].find("PGA") - 3 # Find the starting position of string "PGA"
    end_pga = start_pga + len("PGA") + 3 # Find the ending position of string "PGA"

    # start_inst = f1[check_string].find("INSTANCE_NUMBER") + 1 # Find the starting position of string "INSTANCE_NUMBER"
    # end_inst = start_inst + len("INSTANCE_NUMBER") - 1 # Find the ending position of string "INSTANCE_NUMBER"

    start_snap = f1[check_string].find("SNAP_ID") # Find the starting position of string "snap"
    end_snap = start_snap + 7 # Find the ending position of string "snap"

    sga_list = {}
    pga_list = {}
    max_sga = 0
    max_pga = 0
    # sga_list = []
    # pga_list = []
    # Iterate through the above 2 lines
    for i in range(begin_memory, end_memory):
        # Gets the entire line
        line = f1[i]
        # Slices the line to include only SGAs & PGAs. Replace ',' with '.' (if file is corrupted)
        sga = line[start_sga:end_sga]
        sga = sga.replace(',', '.')
        pga = line[start_pga:end_pga]
        pga = pga.replace(',', '.')
        # Removes extra spaces
        sga = sga.strip()
        pga = pga.strip()
        # Convert from string to float
        sga = float(assignZero(sga))
        pga = float(assignZero(pga))
        # Add SGA and PGA to list
        # sga_list.append(sga)
        # pga_list.append(pga)

        # Do above for snap, add to dictionary as key and give a initial value of 0
        snap = line[start_snap:end_snap]
        snap = snap.replace(',', '.')
        snap = snap.strip()
        snap = int(snap)
        if snap not in sga_list:
            sga_list[snap] = 0
        if snap not in pga_list:
            pga_list[snap] = 0
        # read_iops_list.append(read_iops)
        # write_iops_list.append(write_iops)

        # Add read_iops and write_iops to list for the particular snap ID
        sga_list[snap] += sga
        pga_list[snap] += pga

        # # Slices the line to include only instance
        # instance_number = line[start_inst:end_inst]
        # # Removes extra spaces
        # instance_number = instance_number.strip()
        # # Convert from string to float
        # instance_number = int(instance_number)

        # Append the SGA & PGA to the corresponding key (ServerName)
        # all_sga[server_list[instance_number-1]].append(sga)
        # all_pga[server_list[instance_number-1]].append(pga)

    # Get maximum of read_iops and write_iops grouped by snap
    keymax_sga = max(sga_list.keys(), key=(lambda k: sga_list[k]))
    max_sga = sga_list[keymax_sga]
    keymax_pga = max(pga_list.keys(), key=(lambda k: pga_list[k]))
    max_pga = pga_list[keymax_pga]

    # Iterate through list to sort according to number of instances
    # max_sga = []
    # max_pga = []

    # Read every i'th object in list
    # for i in range(instances):
    #     max_sga.append(normal_round(max(sga_list[::(i+1)])))
    #     print("Max SGA for instance %s: %d" % (server_list[i], max_sga[i]))
    #     max_pga.append(normal_round(max(pga_list[::(i+1)])))
    #     print("Max PGA for instance %s: %d" % (server_list[i], max_pga[i]))

    # # Divide list by the number of instances
    # max_sga = [normal_round(i/instances) for i in max_sga]
    # max_pga = [normal_round(i/instances) for i in max_pga]

    # # Print max_sga & max_pga for every instance
    # for i in range(instances):
    #     print("max_sga for instance %s: %d" % (server_list[i], max_sga[i]))
    #     print("max_pga for instance %s: %d" % (server_list[i], max_pga[i]))

    # insert all values to CSV file
    for i in range(instances):
        # insertToCSV(databaseName, instances,
        #             server_list[i], IOGrowth, max(all_read_iops[server_list[i]]), max(all_write_iops[server_list[i]]), max(all_sga[server_list[i]]), max(all_pga[server_list[i]]))
        # insertToCSV(databaseName, instances,
        #             server_list[i], IOGrowth, normal_round((max(all_read_iops))/instances), normal_round((max(all_write_iops))/instances), normal_round(max(all_sga[server_list[i]])), normal_round(max(all_pga[server_list[i]])))
        insertToCSV(databaseName, instances, server_list[i], IOGrowth, normal_round(max_read_iops/instances), normal_round(max_write_iops/instances), normal_round(max_sga/instances), normal_round(max_pga/instances))


'''
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
'''
# Author - Nitin Bhagat

# this section should validate the mappings

def validateMappingSheet():
    try:
        print("Validating inputs.... Please click on the excel dialog box")
        xlApp = win32com.client.DispatchEx('Excel.Application')
        xlsPath = os.path.expanduser(
            workbookFile)
        wb = xlApp.Workbooks.Open(Filename=xlsPath)
        xlApp.Run('Validate')
        wb.Save()
        xlApp.Quit()
        print("Macro ran successfully!")
    except Exception as e:
        print(e)
        print("Error found while running the excel macro!")
        xlApp.Quit()

'''
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
############################################################################################################
'''
# Author - Nitin Bhagat

# this section should run the main functions


# MAIN FUNCTION
for outFile in outFiles:
    file = open(stage_directory + '\\' + outFile)
    f1 = file.readlines()
    getAllValues(f1)
    file.close()



wbk = openpyxl.load_workbook(
    filename=workbookFile, read_only=False, keep_vba=True)
wks = wbk['DB<->Server Mappings']
for r in range(1, len(row)):
    print("Inserting Row %d..." % r)
    x = r + 2
    wks.cell(row=x, column=3).value = row[r][2]
    wks.cell(row=x, column=4).value = row[r][3]
    wks.cell(row=x, column=5).value = row[r][4]
    wks.cell(row=x, column=6).value = row[r][5]
    wks.cell(row=x, column=11).value = row[r][6]
    wks.cell(row=x, column=12).value = row[r][7]



wbk.save(workbookFile)
wbk.close

print("validate mapping sheet 3")
validateMappingSheet()

# Remove all entries under Read Optimization Total I/O and Write Optimization Total I/O under DB <-> Server Mappings
wbk = openpyxl.load_workbook(
    filename=workbookFile, read_only=False, keep_vba=True)
wks = wbk['DB<->Server Mappings']
print("Removing values under column Read Optimization Total I/O & Write Optimization Total I/O...")
for r in range(1, len(row)):
    x = r + 2
    wks.cell(row=x, column=7).value = None
    wks.cell(row=x, column=8).value = None

wbk.save(workbookFile)
wbk.close

print("Script end!")
print("ReadServerInputs.py DONE")
