import openpyxl
import os
import sys

def getMHCfilename():
    filename = input("Enter MHC filename: ")
    return filename.lower()

# Get filepath to folder containing model files
def getFilepathRootFolder():
    filepath = input("Enter filepath to MHC file: ")
    return filepath

# filepath to target file inside folder
import os

def getFilepathTargetFile(filePathRootFolder, filename, fileNumber, extension):
    return os.path.join(filePathRootFolder, f"{filename}{fileNumber}.{extension.upper()}")

def getFilepathTargetFile_MHC(filePathRootFolder, filename, extension  = "mhc"):
    return os.path.join(filePathRootFolder, f"{filename}.{extension.upper()}")

def getNumBranches(filepathRootFolder,MHCfilename):
    BranchFilePath = getFilepathTargetFile(filepathRootFolder, MHCfilename,"", "brn")
    try:
        with open(BranchFilePath) as file:
            lines = file.readlines()
            if len(lines) > 1:
                return int(lines[1].strip())
    except FileNotFoundError:
        print(f"Error: Branch file not found at {BranchFilePath}")
    except ValueError:
        print(f"Error: Could not parse number of branches in {BranchFilePath}")
    return 0

# Removes quotation mark or whitespace from node-label inside inv file
def labelFormat(label):
    return label.strip(' "')

# adds new branches to meta-list
def addBranches(metaList):
    MHCfilename = getMHCfilename()
    filepathRootFolder = getFilepathRootFolder()

    while True:
        print("\nQuantify\n1.Finish\n2.Modelled Network\n3.Customize Network")
        userInput = input("Enter: ")

        if userInput == "1":
            break
        elif userInput == "2":
            totalBranches = getNumBranches(filepathRootFolder, MHCfilename)
            for branch_idx in range(1, totalBranches + 1):
                fNumStr = f"{branch_idx:03d}"
                filepath_inv = getFilepathTargetFile(filepathRootFolder, MHCfilename, fNumStr, "inv")
                
                try:
                    with open(filepath_inv) as file:
                        lines = file.readlines()
                        # Skip header (first 3 lines)
                        data_lines = [l for l in lines[3:] if l.strip()]
                        if not data_lines:
                            continue
                        
                        # First node
                        first_row = data_lines[0].split(",")
                        start_node = labelFormat(first_row[2] if len(first_row) > 2 else first_row[0])
                        
                        # Last node
                        last_row = data_lines[-1].split(",")
                        end_node = labelFormat(last_row[2] if len(last_row) > 2 else last_row[0])
                        
                        metaList.append([start_node, end_node])
                except FileNotFoundError:
                    print(f"Error: INV file not found at {filepath_inv}")
                    sys.exit(1)
        elif userInput == "3":
            upstream = input("Upstream node label: ").strip()
            downstream = input("Downstream node label: ").strip()
            if upstream and downstream:
                metaList.append([upstream, downstream])
        else:
            print("Invalid input. Please choose 1, 2, or 3.")

    return MHCfilename, filepathRootFolder

def getNGL_MHC(filepath, node):
    mhc_node = ""
    try:
        with open(filepath) as file:
            rowCountMHC = 1
            for line in file:
                if rowCountMHC > 13:
                    row = line.strip().split(",")
                    try:
                        mhc_node = labelFormat(row[0])
                    except IndexError:
                        print("Error: row not delimited by comma in MHC file")

                    row = line.strip()
                    columnCount = 1
                    trackReached = 0
                    reachedNewColumn = False
                    endOfNodeLabel = False
                    nglString = ""
                    for char in row:
                        
                        if not endOfNodeLabel:
                            if char == ",":
                                endOfNodeLabel = True

                        if endOfNodeLabel and char != ",":
                            if char != " ":
                                if trackReached == 1:
                                    reachedNewColumn = False
                                else:
                                    reachedNewColumn = True
                                trackReached = 1
                            else:
                                trackReached = 0

                        if reachedNewColumn:
                            columnCount += 1
                            reachedNewColumn = False
                        
                        if columnCount == 4:
                            nglString += char

                rowCountMHC += 1

                if mhc_node.lower() == node.lower():
                    return float(nglString)

    except FileNotFoundError:
        print(f"Error opening MHC file at: {filepath}")
        return None

# Removes quotation marks from pipe type if any but keeps whitespace between words
def pipeTypeFormat(pipeType):
    return pipeType.strip(' "')

# Retrieve data e.g) ".INV"  for a given list of branches
def transferData(branches,MHCfilename,filepathRootFolder):
    #MHCfilename = getMHCfilename()
    numberOfBranches = len(branches)
    nodeLabel = []
    nodeDiameter = []
    nodePipeType = []
    fileNumbersFound = [] # tracking list of filenumbers found as nodes for custom branch are being found 
    fileNumbersEachNode = [] # list of filenumber assoiciated with any node found for every branch
    adjustedfileNumbersFound = [] # list of non-repeating filenumbers from 'fileNumbersEachNode' list
    fileNumbers_slopes = [] 
    chainages = []
    chainagesDrops_NodeTransitions = []
    invertLevels = []
    slopes = []
    NGLeachNode = []
    everyNGL = [] # to contain ngl values such that each branch(longsection) is stored in a separate list from the rest of the branches and that each branch contains a list of ngl values retrieved from upstream node to downstream node
    everyChainage = []
    everyIL = []
    #filepathRootFolder = getFilepathRootFolder()

    # todo: flag to indicate if node is not found in any inv file - to avoid infinite loop
    # Retrieve Node labels, chainages, inner diameter and pipe type for each branch from INV files
    for branch in range(numberOfBranches):

        upstreamNode = branches[branch][0]
        downstreamNode = branches[branch][1]
        tempNodeLabel = []
        tempNodeChainages = []
        tempNodeFilenumbers = []
        tempNodeDiameter = []
        tempNodePipeType = []
               
        proceed = True
        fileNumber = 1
        skipFileNumber = 0
        nodeID = ""
        upstreamNodeFound = False
        
        while (proceed): 
                
                if upstreamNodeFound:
                    if fileNumber == skipFileNumber:
                        fileNumber += 1

                if fileNumber < 10:  
                    fNumberStr = "00" + str(fileNumber)
                else:
                    fNumberStr = "0" + str(fileNumber)

                filepath = getFilepathTargetFile(filepathRootFolder, MHCfilename, fNumberStr, "inv")
                
                try:
                    
                    with open(filepath) as file:
                        rowCount = 1
                        upstreamNodeFound = False
            
                        for line in file:
                            
                            if rowCount > 3:
                                row = line.strip().split(",")
                                try:
                                    nodeID = labelFormat(row[2])
                                    innerDia = float(row[3])
                                    pipeType = pipeTypeFormat(row[6])
                                except ValueError:
                                    row = line.strip().split("  ")
                                    nodeID = labelFormat(row[2])
                                    innerDia = float(row[3])
                                    pipeType = pipeTypeFormat(row[6])

                                if nodeID.lower() == upstreamNode.lower():
                                    upstreamNodeFound = True
                                    skipFileNumber = fileNumber
                                    fileNumbersFound.append(fNumberStr)
                                
                                # resolves duplication of data (drop manholes) found from model files 
                                if upstreamNodeFound and proceed:
                                    if len(tempNodeLabel) == 0:
                                        tempNodeLabel.append(nodeID)
                                        tempNodeChainages.append(float(row[0]))
                                        tempNodeFilenumbers.append(fNumberStr)
                                        tempNodeDiameter.append(innerDia)
                                        tempNodePipeType.append(pipeType)

                                    else:
                                        # resolve duplication of node labels found
                                        lastNodeLabelAppended = tempNodeLabel[-1]
                                        if lastNodeLabelAppended.lower() == nodeID.lower():
                                            tempNodeLabel[-1] = nodeID
                                            tempNodeChainages[-1] = float(row[0])
                                            tempNodeFilenumbers[-1] = fNumberStr

                                            if len(tempNodeDiameter) == 0:
                                                tempNodeDiameter.append(innerDia)
                                                tempNodePipeType.append(pipeType)
                                            else:
                                                tempNodeDiameter[-1] = innerDia
                                                tempNodePipeType[-1] = pipeType

                                        else:
                                            tempNodeLabel.append(nodeID)
                                            tempNodeChainages.append(float(row[0]))
                                            tempNodeFilenumbers.append(fNumberStr)
                                            tempNodeDiameter.append(innerDia)
                                            tempNodePipeType.append(pipeType)

                                # if the last node in the branch is found
                                if nodeID.lower() == downstreamNode.lower() and upstreamNodeFound:
                                    lastNodeLabelAppended = tempNodeLabel[-1]
                                    if lastNodeLabelAppended.lower() == nodeID.lower():
                                        tempNodeLabel[-1] = nodeID
                                        tempNodeChainages[-1] = float(row[0])
                                        tempNodeFilenumbers[-1] = fNumberStr

                                        if len(tempNodeDiameter) == 0:
                                            tempNodeDiameter.append(innerDia)
                                            tempNodePipeType.append(pipeType)
                                        else:
                                            tempNodeDiameter[-1] = innerDia
                                            tempNodePipeType[-1] = pipeType

                                    else:
                                        tempNodeLabel.append(nodeID)
                                        tempNodeChainages.append(float(row[0]))
                                        tempNodeFilenumbers.append(fNumberStr)
                                        tempNodeDiameter.append(innerDia)
                                        tempNodePipeType.append(pipeType)
        
                                    proceed = False

                            rowCount += 1

                        # updates the "upstream node" to the next node in the branch
                        if upstreamNodeFound and proceed:
                            upstreamNode = nodeID
                            fileNumber = 1
                        else:
                            fileNumber += 1
                        
                except FileNotFoundError:
                    print(f"Error: INV file not found at {filepath}")
                    sys.exit(1)
    #END - Retrieve Node labels, chainages, inner diameter and pipe type for each branch from INV files

        nodeLabel.append(tempNodeLabel)
        chainages.append(tempNodeChainages)
        nodeDiameter.append(tempNodeDiameter)
        nodePipeType.append(tempNodePipeType)
        fileNumbersEachNode.append(tempNodeFilenumbers)

    lenFileNumbersFound = len(fileNumbersFound)
    if lenFileNumbersFound > 0:
        adjustedfileNumbersFound.append(fileNumbersFound[0])
        # append unique filenumbers
        for i in range(lenFileNumbersFound - 1):
            if fileNumbersFound[i+1] != fileNumbersFound[i]:
                adjustedfileNumbersFound.append(fileNumbersFound[i+1])
    #END - Retrieve Node labels and chainages for each branch from INV files
   
    # Retrieve chainages & at drops, node transitions as well as invert levels only at the nodes
    outerLenNodeLabel= len(nodeLabel)
    for outerIndex in range(outerLenNodeLabel):
        innerLenNodeLabel = len(nodeLabel[outerIndex])
        fileNumPrevStr = ""
        fileNumPrev = 0
        fileNumCurrent = 0
        tempInvertLevels = []
        tempChainagesDrops_NodeTransitions = []
        tempFileNumbers_slopes = []

        for innerIndex in range(innerLenNodeLabel):
            fileNumChanged = False
            if innerIndex > 0:
                fileNumPrev = fileNumCurrent
                fileNumPrevStr = fileNumStr

            nodeID = nodeLabel[outerIndex][innerIndex]
            fileNumStr = fileNumbersEachNode[outerIndex][innerIndex]
            fileNumCurrent = int(fileNumStr)
            filepathIL = getFilepathTargetFile(filepathRootFolder, MHCfilename, fileNumStr, "inv")
            foundNodeTwice = False
            foundNodeAtleastOnce = False
            
            if innerIndex > 0:
                if abs(fileNumPrev - fileNumCurrent) > 0:
                    fileNumChanged = True

            try:
                with open(filepathIL) as file:
                    rowCount_IL = 1
                    nodeIL = 0
                    for line in file:
                        if rowCount_IL > 3:
                            row = line.strip().split(",")
                            try:
                                if rowCount_IL > 4:
                                    nodeIL_Prev = nodeIL
                                nodeChainage = float(row[0])
                                nodeIL = float(row[1])
                                nodeLabel_INV = labelFormat(row[2])
                            except ValueError:
                                row = line.strip().split("  ")
                                nodeChainage = float(row[0])
                                nodeIL = float(row[1])
                                nodeLabel_INV = labelFormat(row[2])
                                 
                            if nodeID.lower() == nodeLabel_INV.lower() and foundNodeTwice == False:
                                if abs(round(nodeIL,2)) == 0.05:
                                    nodeIL_Prev -= 0.05
                                    nodeIL_Prev = round(nodeIL_Prev,3)
                                    foundNodeTwice = True
                                    tempInvertLevels.append(nodeIL_Prev)
                                    tempChainagesDrops_NodeTransitions.append(nodeChainage)
                                    tempFileNumbers_slopes.append(fileNumStr)
                                    break
                                else:
                                    tempInvertLevels.append(round(nodeIL,3))
                                    if fileNumChanged:
                                        chainage_LinkingNode = 0
                                        filepath_PrevFile = getFilepathTargetFile(filepathRootFolder, MHCfilename, fileNumPrevStr, "inv")
                                        try:
                                            with open(filepath_PrevFile) as file:
                                                rowCount_LinkingNode = 1
                                                for line in file:
                                                    if rowCount_LinkingNode > 3:
                                                        row = line.strip().split(",")

                                                        try:
                                                            nodeLabel_LinkingNode = labelFormat(row[2])
                                                        except ValueError:
                                                            row = line.strip().split("  ")
                                                            nodeLabel_LinkingNode = labelFormat(row[2])

                                                        if nodeID.lower() == nodeLabel_LinkingNode.lower():
                                                            chainage_LinkingNode = float(row[0])
                                                            break

                                                    rowCount_LinkingNode += 1

                                        except FileNotFoundError:
                                            print(f"Error: Linking node file not found at {filepath_PrevFile}")
                                            sys.exit(1)
                                        
                                        tempChainagesDrops_NodeTransitions.append(chainage_LinkingNode)
                                        tempFileNumbers_slopes.append(fileNumPrevStr)
                                        foundNodeAtleastOnce = True

                                    else:
                                        tempChainagesDrops_NodeTransitions.append(nodeChainage)
                                        tempFileNumbers_slopes.append(fileNumStr)
                                        foundNodeAtleastOnce = True

                            if foundNodeAtleastOnce:
                                if nodeID.lower() != nodeLabel_INV.lower():
                                    break

            except FileNotFoundError:
                print(f"Error: INV file not found at {filepathIL}")
                sys.exit(1)

        invertLevels.append(tempInvertLevels)
        chainagesDrops_NodeTransitions.append(tempChainagesDrops_NodeTransitions)
        fileNumbers_slopes.append(tempFileNumbers_slopes)
    #END - Retrieve chainages + at drops and node transitions as well as invert levels only at the nodes

    # Calculate slopes between nodes
    outerLen_invert = len(invertLevels)
    for outerIndex_invert in range(outerLen_invert):
        tempSlopes = []
        innerLen_invert = len(invertLevels[outerIndex_invert])
        for innerIndex_invert in range(innerLen_invert - 1):
            IL_In = invertLevels[outerIndex_invert][innerIndex_invert]
            IL_Out = invertLevels[outerIndex_invert][innerIndex_invert + 1]
            Chainage_In = chainagesDrops_NodeTransitions[outerIndex_invert][innerIndex_invert]
            Chainage_Out = chainagesDrops_NodeTransitions[outerIndex_invert][innerIndex_invert + 1]
            fileNum_In = int(fileNumbers_slopes[outerIndex_invert][innerIndex_invert])
            fileNum_Out = int(fileNumbers_slopes[outerIndex_invert][innerIndex_invert + 1])
            fileNumChanged_slopes = False

            if abs(fileNum_In - fileNum_Out) > 0:
                fileNumChanged_slopes = True
            
            if abs(Chainage_Out - Chainage_In) and fileNumChanged_slopes == False:
                slope = abs((IL_Out - IL_In) / (Chainage_Out - Chainage_In))
                tempSlopes.append(slope)
        
        slopes.append(tempSlopes)
    #END - Calculate slopes between nodes

    # Retrieve NGL for nodes & intermediate points and calculate inverts at these points
    # if NGL of any node in branch is not found, retrieve NGL value from MHC file and proceed
    lenChainages = len(chainages)
    for outerListStep in range(lenChainages):
        lenInnerChainages = len(chainages[outerListStep])   
        fileNumPrev = 0
        fileNumCurrent = 0
        rowCountCache = 1000000
        filepathNGL = ""
        node = ""
        chainageIndex_atNode = -1
        newInvert = 0
        adjustIndex_accessIL = 0
        adjustIndex_accessDrop = 0
        ngl_foundInMHC = False
        Inverts_Branch = []
        NGLs_Branch = []
        Chainages_Branch = []

        for innerListStep in range(lenInnerChainages):
            # Determine change of filenumber
            fileNumChanged = False
            if innerListStep > 0:
                fileNumPrev = fileNumCurrent
                # pad filenumber with leading zeros
                if fileNumPrev < 10:
                    fileNumPrevStr = "00" + str(fileNumPrev)
                elif fileNumPrev > 9 and fileNumPrev < 100:
                    fileNumPrevStr = "0" + str(fileNumPrev)
                else:
                    fileNumPrevStr = str(fileNumPrev)

                filepathINV = getFilepathTargetFile(filepathRootFolder, MHCfilename, fileNumPrevStr, "inv")

            xPrev = 0
            xCurrent = 0
            delta_x = 0
            segmentSlope = 0
            
            if innerListStep > 0:
                segmentSlope = slopes[outerListStep][chainageIndex_atNode]

            IL_fromList_current = invertLevels[outerListStep][innerListStep] # TBD ?
            tempEveryNGL = []
            tempEveryChainage = []
            tempEveryIL = []
            tempNGLeachNode = []
            fNumberStrNGL = fileNumbersEachNode[outerListStep][innerListStep]
            fileNumCurrent = int(fNumberStrNGL)
            foundNodeChainage = chainages[outerListStep][innerListStep]
            chainageOut = foundNodeChainage
            node = nodeLabel[outerListStep][innerListStep]
            lenIL_Branch = len(invertLevels[outerListStep])
            filepathPrev = filepathNGL
            filepathNGL = getFilepathTargetFile(filepathRootFolder, MHCfilename, fNumberStrNGL, "ngl")

            # get chainage from inv file of linking node - (chainage-in)
            if innerListStep > 0:
                if abs(fileNumCurrent - fileNumPrev) != 0:
                    fileNumChanged = True

                    except FileNotFoundError:
                        print(f"Error: INV file not found at {filepathINV}")
                        sys.exit(1)
             
            try:
                if fileNumChanged:
                    masterFilePath = filepathPrev
                else:
                    masterFilePath = filepathNGL

                with open(masterFilePath) as file:
                    rowCount = 1
                     
                    isNodeChainage = True 

                    for line in file:

                        if rowCount > 4:
                            if rowCount > 5:
                                nglFileChainagePrev = nglFileChainage

                            row = line.strip().split("  ")
                            try:
                                nglFileChainage = float(row[0])
                            except ValueError:
                                row = line.strip().split(",")
                                nglFileChainage = float(row[0])

                            # if NGL of any node in branch is not found, retrieve NGL value from MHC file and proceed
                            if round((nglFileChainage - foundNodeChainage),3) > 0.001 and isNodeChainage:
                                mhcFilepath = getFilepathTargetFile_MHC(filepathRootFolder, MHCfilename)
                                ngl_MHC = getNGL_MHC(mhcFilepath, node)
                                if ngl_MHC != None:
                                    ngl_foundInMHC = True

                            if abs(round(nglFileChainage - foundNodeChainage,3)) <= 0.001 and isNodeChainage or ngl_foundInMHC:
                                # todo: if ngl_foundInMHC is true, append ngl_MHC to tempEveryNGL
                                if ngl_foundInMHC:
                                    tempEveryNGL.append(ngl_MHC)
                                    tempEveryChainage.append(foundNodeChainage)
                                    ngl_foundInMHC = False
                                else:
                                    tempEveryNGL.append(float(row[1]))
                                    tempEveryChainage.append(nglFileChainage)
                                
                                # retrieve invert level at node. if drop at node, append drop IL else append IL without drop
                                if innerListStep > 0:
                                    adjustIndex_accessDrop += 1
                                    if innerListStep + adjustIndex_accessDrop < lenIL_Branch:
                                        IL_fromList_next = invertLevels[outerListStep][innerListStep + adjustIndex_accessDrop]
                                    if innerListStep + adjustIndex_accessIL < lenIL_Branch:
                                        IL_fromList_current = invertLevels[outerListStep][innerListStep + adjustIndex_accessIL]
                                        tempEveryIL.append(IL_fromList_current)

                                    if round(IL_fromList_current - IL_fromList_next, 2) == 0.05:
                                        adjustIndex_accessIL += 1
                                        newInvert = IL_fromList_next
                                    else:
                                        adjustIndex_accessDrop -= 1
                                        newInvert = IL_fromList_current

                                else:
                                    IL_fromList_current = invertLevels[outerListStep][innerListStep]
                                    tempEveryIL.append(IL_fromList_current)
                                    newInvert = IL_fromList_current

                                isNodeChainage = False
                                rowCountCache = rowCount
                                chainageIndex_atNode = innerListStep

                                break
                            

                            if fileNumChanged == False and rowCount > rowCountCache:
                                tempEveryNGL.append(float(row[1]))
                                tempEveryChainage.append(nglFileChainage)
                                rowCountCache = rowCount

                                if rowCount > 5:
                                    xPrev = nglFileChainagePrev
                                    xCurrent = nglFileChainage
                                    delta_x = xCurrent - xPrev
                                    newInvert = newInvert - (segmentSlope * delta_x)
                                    tempEveryIL.append(round(newInvert, 3))

                            if fileNumChanged and rowCount > rowCountCache:
                                tempEveryNGL.append(float(row[1]))
                                tempEveryChainage.append(nglFileChainage)
                                rowCountCache = rowCount

                                if rowCount > 5:
                                    xPrev = nglFileChainagePrev
                                    xCurrent = nglFileChainage
                                    delta_x = xCurrent - xPrev
                                    newInvert = newInvert - (segmentSlope * delta_x)
                                    tempEveryIL.append(round(newInvert, 3))
                        rowCount += 1

                # opens ngl file associated with "chainageOut" / linking node to determine the row count found at "chainageOut"(of the linking node) so that "rowCountCache" is updated - to know where to continue reading from in the next ngl file if filenumber has changed
                if fileNumChanged:
                    try:
                        with open(filepathNGL) as file:
                            rowCount_fileChanged = 1
                            for line in file:
                                if rowCount_fileChanged > 4:
                                    row = line.strip().split("  ")
                                    try:
                                        nglFileChainage_fileChanged = float(row[0])
                                    except ValueError:
                                        row = line.strip().split(",")
                                        nglFileChainage_fileChanged = float(row[0])

                                    if abs(round(nglFileChainage_fileChanged - chainageOut,3)) <= 0.001:
                                        rowCountCache = rowCount_fileChanged
                                        break

                                rowCount_fileChanged += 1                       

                    except FileNotFoundError:
                        print(f"Error: NGL file not found at {filepathNGL}")
                        sys.exit(1)

            except FileNotFoundError:
                print(f"Error: Master file not found at {masterFilePath}")
                sys.exit(1)
    #END - Retrieve NGL for nodes & intermediate points and calculate inverts at these points

            # Append data for each node in branch 
            NGLeachNode.append(tempNGLeachNode)
            NGLs_Branch.append(tempEveryNGL)
            Chainages_Branch.append(tempEveryChainage)
            Inverts_Branch.append(tempEveryIL)
            
        # Append data for each branch in meta-list
        everyNGL.append(NGLs_Branch)
        everyChainage.append(Chainages_Branch) 
        everyIL.append(Inverts_Branch)

    return nodeLabel, everyChainage, everyNGL, everyIL, nodeDiameter, nodePipeType, slopes

def OutsideDiameter_Sewer(pipeType):

    OD = []
    numberOfBranches = len(pipeType)

    for branch in range(numberOfBranches):
        branchLength = len(pipeType[branch])
        tempOD = []

        for element in range(branchLength):
            pipeSegment = pipeType[branch][element]
            splitPipeSegment = pipeSegment.split(" ")
            outsideDiameter = splitPipeSegment[0]
            secondSplit = outsideDiameter.split("mm")
            outsideDiameter = float(secondSplit[0])
            tempOD.append(outsideDiameter)

        OD.append(tempOD)

    return OD

def generateSpreadsheet(nodeLabels, pipeTypes, innerDiameters, outsideDiameters, chainages, NGLs, ILs):
    rowNumberTotals = 0
    rowNumber_excavations = 0
    rowNumber_hardRock = 0
    rowNumber_granularFill_Reuse = 0
    rowNumber_selectedFill_Reuse = 0
    rowNumber_granularFill_Import = 0
    rowNumber_selectedFill_Import = 0
    rowNumber_totalBackfilling = 0
    rowNumber_importedBackfillMaterial = 0
    rowNumber_excessMaterial = 0

    HEADERS = ["Node ID", "Pipe Type", "Inner Diameter (m)", "Outside Diameter (m)", "Bedding Depth(m)", "Pipe Thickeness (m)", "Working Space (m)", " ", "Chainage (m)", "Distance (m)","NGL (m)", "IL (m)", "Trench Level (m)", "Trench Depth (m)", "Trench Width (m)", "Excavation (m\u00B3)", " ", "0-1m Deep Excavation (m\u00B3)", "1-2m Deep Excavation (m\u00B3)", "2-3m Deep Excavation (m\u00B3)", "3-4m Deep Excavation (m\u00B3)", "4-5m Deep Excavation (m\u00B3)", "5m-6m Deep Excavation (m\u00B3)",">6m Deep Excavation (m\u00B3)", " ", "0-1m Deep Excavation (m\u00B3)", "1-2m Deep Excavation (m\u00B3)", "2-3m Deep Excavation (m\u00B3)", "3-4m Deep Excavation (m\u00B3)", "4-5m Deep Excavation (m\u00B3)", "5-6m Deep Excavation (m\u00B3)", ">6m Deep Excavation (m\u00B3)" ," ", "Bedding Backfill (m\u00B3)", "Selected Backfill (m\u00B3)", "Additional Selected Backfill (m\u00B3)", "Backfill (m\u00B3)"]
    Summary_Metrics = ["Excavations", "Hard Rock", " ", "Granular Fill (Reuse)", "Selected Fill (Reuse)", "Granular Fill (Import)", "Selected Fill (Import)", " ", "Total Backfilling", "Imported Backfill Material", "Excess Material"]
    metrics_Percentages = [0, 0.2, 0, 0.3, 0.3, 0.7, 0.7, 0, 0, 1, 0]

    #Summary_Metrics_Formulas = [f"=P{rowNumberTotals}", f"=S{rowNumber_excavations} * T{rowNumber_hardRock}", " ", f"=AH{rowNumberTotals} * T{rowNumber_granularFill_Reuse}", f"=AI{rowNumberTotals} * T{rowNumber_selectedFill_Reuse}", f"=AH{rowNumberTotals} * T{rowNumber_granularFill_Import}", f"=AI{rowNumberTotals} * T{rowNumber_selectedFill_Import}", " ", f"=AJ{rowNumberTotals} + AK{rowNumberTotals}", f"=S{rowNumber_totalBackfilling} - (S{rowNumber_excavations} - S{rowNumber_hardRock} * T{rowNumber_importedBackfillMaterial} - S{rowNumber_granularFill_Reuse} - S{rowNumber_selectedFill_Reuse})", f"=S{rowNumber_hardRock}"]
    header_alignment = openpyxl.styles.Alignment(wrap_text=True, horizontal='center', vertical='center')
    header_font = openpyxl.styles.Font(bold=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    data_write_start_row = -1

    for col_idx, header in enumerate(HEADERS, start=1):
        ws.cell(row=2, column=col_idx, value=header)
        ws.cell(row=2, column=col_idx).alignment = header_alignment
        cell = ws.cell(row=2, column=col_idx)
        cell.font = header_font
        column_letter = openpyxl.utils.get_column_letter(col_idx)

        if header == " ":
            ws.column_dimensions[column_letter].width = 2
        elif header == "Pipe Type":
            ws.column_dimensions[column_letter].width = 25
        elif header == "0-1m Deep Excavation (m\u00B3)":
            ws.column_dimensions[column_letter].width = 25
        else:
            ws.column_dimensions[column_letter].width = 10

    numberOfBranches = len(chainages)
    detectBranchChange = False
    trackBranchIndex = 0
    skipRowAfterBranchChange = []
    for branchIndex in range(numberOfBranches):
        numberOfSegments = len(chainages[branchIndex])
        if branchIndex > 0:
            if trackBranchIndex - branchIndex != 0:
                detectBranchChange = True

        for segmentIndex in range(numberOfSegments):
            numberOfPoints = len(chainages[branchIndex][segmentIndex])

            for pointIndex in range(numberOfPoints):
                if detectBranchChange:
                    row = ws.max_row + 2
                    detectBranchChange = False
                    skipRowAfterBranchChange.append(row - 1)
                else:
                    row = ws.max_row + 1

                if data_write_start_row == -1:
                    data_write_start_row = row

                if pointIndex + 1 == numberOfPoints:
                    ws.cell(row=row, column=1, value=nodeLabels[branchIndex][segmentIndex])
                
                ws.cell(row=row, column=2, value=pipeTypes[branchIndex][segmentIndex])
                ws.cell(row=row, column=3, value=innerDiameters[branchIndex][segmentIndex] / 1000)
                ws.cell(row=row, column=4, value=outsideDiameters[branchIndex][segmentIndex] / 1000)
                ws.cell(row=row, column=5, value=f"=IF((D{row}/4)>0.2,0.2, IF((D{row}/4)<0.1,0.1, (D{row}/4)))")
                ws.cell(row=row, column=6, value=f"=(D{row} - C{row})/2")
                ws.cell(row=row, column=7, value=0.3)

                ws.cell(row=row, column=9, value=chainages[branchIndex][segmentIndex][pointIndex])
                ws.cell(row=row, column=10, value=f"=IFERROR(I{row}-I{row-1}, 0)")
                ws.cell(row=row, column=11, value=NGLs[branchIndex][segmentIndex][pointIndex])
                ws.cell(row=row, column=12, value=ILs[branchIndex][segmentIndex][pointIndex])
                ws.cell(row=row, column=13, value=f"=L{row} - F{row} - E{row}")
                ws.cell(row=row, column=14, value=f"=K{row} - M{row}")
                ws.cell(row=row, column=15, value=f"=D{row} + 2 * G{row}")
                ws.cell(row=row, column=16, value=f"=IFERROR(J{row} * O{row} * (N{row} + N{row-1}) / 2, 0)")
                ws.cell(row=row, column=18, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 < 1, (O{row} < 1)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")

                ws.cell(row=row, column=19, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 1, (N{row} + N{row-1}) / 2 < 2, (O{row} < 1)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=20, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 2, (N{row} + N{row-1}) / 2 < 3, (O{row} < 1)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=21, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 3, (N{row} + N{row-1}) / 2 < 4, (O{row} < 1)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=22, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 4, (N{row} + N{row-1}) / 2 < 5, (O{row} < 1)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=23, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 5, (N{row} + N{row-1}) / 2 < 6, (O{row} < 1)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=24, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 6, (O{row} < 1)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")

                ws.cell(row=row, column=26, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 < 1, (O{row} >= 1), (O{row} < 2)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=27, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 1, (N{row} + N{row-1}) / 2 < 2, (O{row} >= 1), (O{row} < 2)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=28, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 2, (N{row} + N{row-1}) / 2 < 3, (O{row} >= 1), (O{row} < 2)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=29, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 3, (N{row} + N{row-1}) / 2 < 4, (O{row} >= 1), (O{row} < 2)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=30, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 4, (N{row} + N{row-1}) / 2 < 5, (O{row} >= 1), (O{row} < 2)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=31, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 5, (N{row} + N{row-1}) / 2 < 6, (O{row} >= 1), (O{row} < 2)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")
                ws.cell(row=row, column=32, value=f"=IFERROR(IF(AND((N{row} + N{row-1}) / 2 >= 6, (O{row} >= 1), (O{row} < 2)), J{row} * O{row} * (N{row} + N{row-1}) / 2, 0), 0)")

                ws.cell(row=row, column=34, value=f"=(J{row} * O{row} * (D{row} + E{row} + F{row}) - PI() * ( (D{row}/2)^2 ) * J{row})")
                ws.cell(row=row, column=35, value=f"=J{row} * O{row} * 0.2")
                ws.cell(row=row, column=37, value=f"=(P{row} - PI() * ( (D{row}/2)^2 ) * J{row}) - AH{row} - AI{row}")

        trackBranchIndex = branchIndex
    
    # Totals
    data_write_end_row = ws.max_row
    row = ws.max_row + 2
    rowNumberTotals = row
    columnMax = ws.max_column
    SUM_COLUMNS = [10,16,18,19,20,21,22,23,24,26,27,28,29,30,31,32,34,35,36,37]

    for col_idx in range(1, columnMax+1):
        if col_idx in SUM_COLUMNS:
            columnLetter = openpyxl.utils.get_column_letter(col_idx)
            ws.cell(row=row, column=col_idx, value=f"=SUM({columnLetter}{data_write_start_row}:{columnLetter}{data_write_end_row})")
            rowNumber_excavations = row

    row = ws.max_row + 2
    for i, metric in enumerate(Summary_Metrics):
        ws.cell(row=row + i, column=18, value=metric)
        
        if metric == "Excavations":
            rowNumber_excavations = row + i
        elif metric == "Hard Rock":
            rowNumber_hardRock = row + i
        elif metric == "Granular Fill (Reuse)":
            rowNumber_granularFill_Reuse = row + i
        elif metric == "Selected Fill (Reuse)":
            rowNumber_selectedFill_Reuse = row + i
        elif metric == "Granular Fill (Import)":
            rowNumber_granularFill_Import = row + i
        elif metric == "Selected Fill (Import)":
            rowNumber_selectedFill_Import = row + i
        elif metric == "Total Backfilling":
            rowNumber_totalBackfilling = row + i
        elif metric == "Imported Backfill Material":
            rowNumber_importedBackfillMaterial = row + i
        elif metric == "Excess Material":
            rowNumber_excessMaterial = row + i

    Summary_Metrics_Formulas = [f"=P{rowNumberTotals}", f"=S{rowNumber_excavations} * T{rowNumber_hardRock}", " ", f"=AH{rowNumberTotals} * T{rowNumber_granularFill_Reuse}", f"=AI{rowNumberTotals} * T{rowNumber_selectedFill_Reuse}", f"=AH{rowNumberTotals} * T{rowNumber_granularFill_Import}", f"=AI{rowNumberTotals} * T{rowNumber_selectedFill_Import}", " ", f"=AJ{rowNumberTotals} + AK{rowNumberTotals}", f"=S{rowNumber_totalBackfilling} - (S{rowNumber_excavations} - S{rowNumber_hardRock} * T{rowNumber_importedBackfillMaterial} - S{rowNumber_granularFill_Reuse} - S{rowNumber_selectedFill_Reuse})", f"=S{rowNumber_hardRock}"]

    for i, formula in enumerate(Summary_Metrics_Formulas):
        ws.cell(row=row + i, column=19, value=formula)

    # Apply number formatting - 3 decimals
    for rowRange in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in rowRange:
            cell.number_format = '0.000'

    for i, percentage in enumerate(metrics_Percentages):
        row_idx = row + i
        if percentage != 0:
            cell = ws.cell(row=row_idx, column=20, value=percentage)
            cell.number_format = '0.00%'

    #apply borders around all cells with data
    thin_border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin'),
                         right=openpyxl.styles.Side(style='thin'),
                         top=openpyxl.styles.Side(style='thin'),
                         bottom=openpyxl.styles.Side(style='thin'))
    
    empty_header_columns = [8,17,25,33]    
    for row in ws.iter_rows(min_row=2, max_row=data_write_end_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.column not in empty_header_columns and cell.row not in skipRowAfterBranchChange:
                cell.border = thin_border

    totals_border = openpyxl.styles.Border(top=openpyxl.styles.Side(style='thin'), bottom=openpyxl.styles.Side(style='double'))
    for row in ws.iter_rows(min_row=rowNumberTotals, max_row=rowNumberTotals, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.column in SUM_COLUMNS:
                cell.border = totals_border

    for row in ws.iter_rows(min_row=rowNumber_excavations, max_row=rowNumber_excessMaterial, min_col=18, max_col=20):
        for cell in row:
            if cell.value != " " and cell.value is not None:
                cell.border = thin_border

    # apply fill colour to input cells
    input_columns = [2,3,4,7,9,11,12]
    for row in ws.iter_rows(min_row=data_write_start_row, max_row=data_write_end_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.column in input_columns and cell.row not in skipRowAfterBranchChange:
                cell.fill = openpyxl.styles.PatternFill(start_color="FFC7E0BD", end_color="FFC7E0BD", fill_type = "solid")
    
    # todo: apply fill colour to percentage cells in summary section
    for row in ws.iter_rows(min_row=rowNumber_excavations, max_row=rowNumber_excessMaterial, min_col=20, max_col=20):
        for cell in row:
            if cell.value != " " and cell.value is not None:
                cell.fill = openpyxl.styles.PatternFill(start_color="FFC7E0BD", end_color="FFC7E0BD", fill_type = "solid")
    
    # apply thick border around header row
    thick_border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='medium'),
                         right=openpyxl.styles.Side(style='medium'),
                         top=openpyxl.styles.Side(style='medium'),
                         bottom=openpyxl.styles.Side(style='medium'))
    for row in ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=ws.max_column):
        for cell in row:
            if cell.value is not None and cell.value != " ":
                cell.border = thick_border
    
    
    wb.save("Quantified_sewer.xlsx")

def main():
    link = []
    MHC_filename, filepath_rootFolder = addBranches(link)
    print(link)
    print(MHC_filename)
    print(filepath_rootFolder)
    retrievedLabels, retrievedChainages, retrievedNGL, retrievedIL, retrievedInnerDiameters, retrievedPipeTypes, retrievedSlopes = transferData(link, MHC_filename, filepath_rootFolder)
    retrievedOutsideDiameters = OutsideDiameter_Sewer(retrievedPipeTypes)
    print(retrievedSlopes[10])
    #print(retrievedLabels)
    #print(retrievedChainages)
    #print(retrievedNGL)
    #print(retrievedIL)
    # print(retrievedInnerDiameters)
    # print(retrievedPipeTypes)
    # print(retrievedOutsideDiameters)
    generateSpreadsheet(retrievedLabels, retrievedPipeTypes, retrievedInnerDiameters, retrievedOutsideDiameters, retrievedChainages, retrievedNGL, retrievedIL)

if __name__ == "__main__": 
    main()
