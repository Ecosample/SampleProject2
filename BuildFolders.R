

fd <- read.csv(file="C:/Users/mbayly/Desktop/Ecodat_jobs/QAQC/setup/folderset.csv")
fd

main <- "C:/Users/mbayly/Desktop/Ecodat_jobs/QAQC/OldLiveExport"
setwd(main)
for(i in 1:length(unique(fd$module))){
	#dir.create(paste0(main, "/", unique(fd$module)[i]))
	dir.create(as.character(unique(fd$module)[i]))
	setwd(paste0(main, "/", unique(fd$module)[i]))
	thislabels <- as.character(fd[which(fd$module == unique(fd$module)[i]), c("label")])
	
	for(j in 1:length(thislabels)){
		dir.create(thislabels[j])
	}

	setwd(main)
}

#line addition
# i did a line addition here!!!
##########################################
##########################################
# TEST FOR ANY DIFFERENCES

require(XLConnect)
library(qdapTools)

# tracking sheet
# https://docs.google.com/spreadsheets/d/1iIZAi4Yi2Bb5YYK5gcQDsjqCt-UCl6VrQj6UmPoV2ZY/edit#gid=0

module = "Form Angling"
module = "Form Facility Inspection"
module = "Form Electrofishing"
module = "Form Beach Seine"
module = "Form Facility Inspection"


module = "Form Beach Seine"
module = "Form Facility Inspection"
module = "Form Fish Habitat Assessment"
module = "Form Fish Habitat Gravel"
module = "Form Electrofishing"


module = "Form Spawning Survey"


basedir <- paste0("C:/Users/mbayly/Desktop/Ecodat_jobs/QAQC/NewLiveExport/", module)
setwd(basedir)
myreports <- list.files()
error.report <- list()
masterlist <- list()


for(j in 1:length(myreports)){
	setwd(paste0(basedir, "/", myreports[j]))
	
	##########################################################
	# FOR AN XLSX FILE
	if(length(list.files(pattern = ".xlsx"))>0){
		wb = loadWorkbook(list.files(pattern = ".xlsx")[1])
		print(list.files(pattern = ".xlsx")[1]);
		(sizeNewxlsx <- file.info(list.files(pattern = ".xlsx")[1])$size)
		print(sizeNewxlsx)
		thissheets <- getSheets(wb)
			new <- data.frame()
				for(i in 1:length(thissheets)){
					df = readWorksheet(wb, sheet = i, header = TRUE)
					mname <- thissheets[i]
					mrow <- dim(df)[1]
					mcol <- dim(df)[2]
					addrow <- data.frame(mname=mname, mrow=mrow, mcol=mcol)
					new <- rbind(new, addrow)
				}
			print(new)
		setwd(paste0("C:/Users/mbayly/Desktop/Ecodat_jobs/QAQC/OldLiveExport/", module, "/",  myreports[j]))
		print(list.files(pattern = ".xlsx")[1]);
		(sizeOldxlsx <- file.info(list.files(pattern = ".xlsx")[1])$size)
		print(sizeOldxlsx)
		wb = loadWorkbook(list.files(pattern = ".xlsx")[1])
		thissheets <- getSheets(wb)
			old <- data.frame()
				for(i in 1:length(thissheets)){
					df = readWorksheet(wb, sheet = i, header = TRUE)
					mname <- thissheets[i]
					mrow <- dim(df)[1]
					mcol <- dim(df)[2]
					addrow <- data.frame(mname=mname, mrow=mrow, mcol=mcol)
					old <- rbind(old, addrow)
				}
			print(old)
			if(!FALSE %in% c(dim(old)==dim(new))){
				try({
				anydiffs <- new == old		
				if(FALSE %in% anydiffs){
					error.report <- append(error.report, paste0("Errors in Excel export in ", myreports[j])) 
					error.report <- append(error.report, data.frame(old)) 
					error.report <- append(error.report, new) 
				}
				})
				rm(new); rm(old); rm(anydiffs)
			} else {
				error.report <- append(error.report, paste0("Excel file missing sheets in ", myreports[j])) 
			}
			
			if(sizeNewxlsx/sizeOldxlsx > 1.05 | sizeNewxlsx/sizeOldxlsx < 0.95){
				error.report <- append(error.report, paste0("Excel file sizes differ between new and old in ", myreports[j])) 
			}	
			rm(sizeNewxlsx); rm(sizeOldxlsx)
	}
	setwd(paste0(basedir, "/", myreports[j]))
	
	##########################################################
	# FOR A KML FILE
	if(length(list.files(pattern = ".kml"))>0){
		library(maptools)
		tkml <- getKMLcoordinates(kmlfile=list.files(pattern = ".kml")[1], ignoreAltitude=T)
		tkmlNew <- unlist(tkml)
		print(list.files(pattern = ".kml")[1]);
		(sizeNewkml <- file.info(list.files(pattern = ".kml")[1])$size); print(sizeNewkml)

		setwd(paste0("C:/Users/mbayly/Desktop/Ecodat_jobs/QAQC/OldLiveExport/", module, "/",  myreports[j]))
		tkml <- getKMLcoordinates(kmlfile=list.files(pattern = ".kml")[1], ignoreAltitude=T)
		tkmlOld <- unlist(tkml)
		print(list.files(pattern = ".kml")[1]);
		(sizeOldkml <- file.info(list.files(pattern = ".kml")[1])$size); print(sizeOldkml)

		if(length(tkmlOld) == length(tkmlNew)){
			if(FALSE %in% c(tkmlOld==tkmlNew)){
				error.report <- append(error.report, paste0("Waypoints differ in ", myreports[j])) 
			}
		} else {
			error.report <- append(error.report, paste0("Number of waypoints differ in ", myreports[j])) 
		}
	rm(tkmlNew);rm(tkmlOld);rm(tkml)
	
	if(sizeNewkml/sizeOldkml > 1.05 | sizeNewkml/sizeOldkml < 0.95){
				error.report <- append(error.report, paste0("KML file sizes differ between new and old in ", myreports[j])) 
			}	
			rm(sizeNewkml); rm(sizeOldkml)
	}
	setwd(paste0(basedir, "/", myreports[j]))
	
	##########################################################
	# FOR A WORD DOC	
	if(length(list.files(pattern = ".docx"))>0){
		sizeNew <- file.info(list.files(pattern = ".docx")[1])$size
		txtNew <- read_docx(list.files(pattern = ".docx")[1])
		print(list.files(pattern = ".docx")[1]); print(sizeNew)

		setwd(paste0("C:/Users/mbayly/Desktop/Ecodat_jobs/QAQC/OldLiveExport/", module, "/",  myreports[j]))
		sizeOld <- file.info(list.files(pattern = ".docx")[1])$size
		txtOld <- read_docx(list.files(pattern = ".docx")[1])
		print(list.files(pattern = ".docx")[1]); ; print(sizeOld)
				
		if(sizeNew/sizeOld > 1.05 | sizeNew/sizeOld < 0.95){
			error.report <- append(error.report, paste0("Picture appendix file size differ in ", myreports[j]))
		}
		if(length(txtNew) != length(txtOld)){
			error.report <- append(error.report, paste0("Picture appendix content different in ", myreports[j]))
		}	
		rm(sizeNew);rm(sizeOld);rm(txtNew);rm(textOld)
	}
	setwd(paste0(basedir, "/", myreports[j]))	

	##########################################################
	# Export missing from old vs new 
	filesNew = length(dir())
	setwd(paste0("C:/Users/mbayly/Desktop/Ecodat_jobs/QAQC/OldLiveExport/", module, "/",  myreports[j]))
	filesOld = length(dir())
	print(filesNew); print(filesOld)
	
	if(filesNew > filesOld){
		error.report <- append(error.report, paste0("Export missing from 'OLD' live ecodat for ", myreports[j]))
	}
	if(filesNew < filesOld){
		error.report <- append(error.report, paste0("Export missing from 'NEW' live ecodat for ", myreports[j]))
	}
	if(filesNew==0 & filesOld==0){
		error.report <- append(error.report, paste0("Export missing from both NEW & OLD for ", myreports[j]))
	}
}	
	
error.report	
	
	

