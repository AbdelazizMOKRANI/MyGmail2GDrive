// Gmail2GDrive
/**
 * Returns the label with the given name or creates it if not existing.
 */
function getOrCreateLabel(labelName) {
	var label = GmailApp.getUserLabelByName(labelName);
	if (label == null) {
		label = GmailApp.createLabel(labelName);
	}
	return label;
}
/**
 * Recursive function to create and return a complete folder path.
 */
function getOrCreateSubFolder(baseFolder, folderArray) {
	if (folderArray.length == 0) {
		return baseFolder;
	}
	var nextFolderName = folderArray.shift();
	var nextFolder = null;
	var folders = baseFolder.getFolders();
	for (var i = 0; i < folders.length; i++) {
		var folder = folders[i];
		if (folders[i].getName() == nextFolderName) {
			nextFolder = folders[i];
			break;
		}
	}
	if (nextFolder == null) {
		// Folder does not exist - create it.
		nextFolder = baseFolder.createFolder(nextFolderName);
	}
	return getOrCreateSubFolder(nextFolder, folderArray);
}
/**
 * Returns the GDrive folder with the given path.
 */
function getFolderByPath(path) {
	var parts = path.split("/");
	if (parts[0] == '') parts.shift(); // Did path start at root, '/'?
	var folder = DriveApp.getRootFolder();
	for (var i = 0; i < parts.length; i++) {
		var result = folder.getFoldersByName(parts[i]);
		if (result.hasNext()) {
			folder = result.next();
		} else {
			throw new Error("folder not found.");
		}
	}
	return folder;
}
/**
 * Returns the GDrive folder with the given name or creates it if not existing.
 */
function getOrCreateFolder(folderName) {
	var folder;
	try {
		folder = getFolderByPath(folderName);
	} catch (e) {
		var folderArray = folderName.split("/");
		folder = getOrCreateSubFolder(DriveApp.getRootFolder(), folderArray);
	}
	return folder;
}
/**
 * Main function that processes Gmail attachments and stores them in Google Drive.
 * Use this as trigger function for periodic execution.
 */
function Gmail2GDrive() {
	if (!GmailApp) return; // Skip script execution if GMail is currently not available (yes this happens from time to time and triggers spam emails!)
	var config = getGmail2GDriveConfig();
	var label = getOrCreateLabel(config.processedLabel);
	var end, start;
	start = new Date();
	//Logger.log("INFO: Starting mail attachment processing.");
	if (config.globalFilter === undefined) {
		//-in:trash to save the attahcment of the trash
		config.globalFilter = "has:attachment -in:drafts -in:spam";
	}
	for (var ruleIdx = 0; ruleIdx < config.rules.length; ruleIdx++) {
		var rule = config.rules[ruleIdx];
		var gSearchExp = config.globalFilter + " " + rule.filter + " -label:" + config.processedLabel;
		if (config.newerThan != "") {
			gSearchExp += " newer_than:" + config.newerThan;
		}
		var doArchive = rule.archive == true;
		var threads = GmailApp.search(gSearchExp);
		//Logger.log("INFO:   Processing rule: " + gSearchExp);
		for (var threadIdx = 0; threadIdx < threads.length; threadIdx++) {
			var thread = threads[threadIdx];
			end = new Date();
			var runTime = (end.getTime() - start.getTime()) / 1000;
			//Logger.log("INFO:     Processing thread: " + thread.getFirstMessageSubject() + " (runtime: " + runTime + "s/" + config.maxRuntime + "s)");
			if (runTime >= config.maxRuntime) {
				//Logger.log("WARNING: Self terminating script after " + runTime + "s")
				return;
			}
			var messages = thread.getMessages();
			for (var msgIdx = 0; msgIdx < messages.length; msgIdx++) {
				var message = messages[msgIdx];
				//Logger.log("INFO:       Processing message: " + message.getSubject() + " (" + message.getId() + ")");
				var messageDate = message.getDate();
				var attachments = message.getAttachments();
				var messageBody = message.getPlainBody();
                var messageSubject =message.getSubject();
                // The substituted value will be contained in the resultSubject variable: get all string Text except the first string between square brackets and 'fwd:' Text also removing whitespaces
                var resultSubject = messageSubject.replace(/\b(Fwd:)(\s|$)|\[(.*?)\]\s*|^\s+|\s+$|\s+(?=\s)/gi,"");
				for (var attIdx = 0; attIdx < attachments.length; attIdx++) {
					var attachment = attachments[attIdx];
					// Found "> -1" the xls or xlsx 
					if ((attachment.getName().indexOf(".xls") > -1) || (attachment.getName().indexOf(".xlsx") > -1)) {
						//Logger.log("INFO:         Processing attachment: " + attachment.getName());
						try {
							var folder = getOrCreateFolder(rule.folder);
							var file = folder.createFile(attachment);
							// convert xls or xslx file to google spread Sheet
							if (file.getMimeType() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || "application/vnd.ms-excel") {
								if (resultSubject){
                                var f = Drive.Files.insert(
                             
                                {
                                title: resultSubject                               
								}
                                , file.getBlob(), {
									convert: true
								});
								}
                                //if the email without Subject Rename the Google Sheet file as the name of Xls File attachment 
                                else{   
                                var f = Drive.Files.insert(
                                {
                                title: file.getName()
                               
								}
                                , file.getBlob(), {
									convert: true
								});                                
                                }
                                
								var id = f.id
								Drive.Files.update(f, id, file.getBlob(), {
									convert: true,
									removeParents: DriveApp.getRootFolder().getId()
								})
							}
							var ssFile = DriveApp.getFileById(id);
							DriveApp.getFolderById(folder.getId()).addFile(ssFile);
							var ss = SpreadsheetApp.open(ssFile);
							var sheet = ss.getSheets()[0];
							sheet.insertColumnAfter(sheet.getLastColumn());
							SpreadsheetApp.flush();
							var sheet = ss.getSheets()[0];
							var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn() + 1)                            
							var values = range.getValues();
                            
							values[0][sheet.getLastColumn()] = "Search Strategy";
                                                        
                            for (var i = 1; i < values.length; i++) {                          
                            var txt = messageBody; 
                            var start_=0, r = [], level=0;
                            for (var j = 0; j < txt.length; j++) {
                              if (txt.charAt(j) == '(') {
                                if (level === 0) start_=j;
                                ++level;
                              }
                              if (txt.charAt(j) == ')') {
                                 
                                if (level > 0) {
                                        --level;
                                }
                                if (level === 0) {
                                    r.push(txt.substring(start_, j+1));
                                }
                              }
                            }   
		            //Regex for extract the keywords that start with these characters (to,an,su(keywords)) 		    
                            var rx = "\\b(?:ti|ab|su)(?:,(ti|ab|su))*\\(";
                            //Extract from the list all the string text that contains these caracteres (ti,ab,su)
                            var result = r.filter(function(y) { return new RegExp(rx, "i").test(y); }).map(function(x) {return x.replace(new RegExp(rx, "ig"), '(')});
                            //in case we have ),( Replace the comma with whitespace 
                            var finalResult=result.toString().replace(/[)],[(]/g,") (");
                            if(finalResult) {values[i][values[i].length - 1] =finalResult;}
                                                       
                            }
                            
						    range.setValues(values);
                            
                            var range2 = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn() + 1)                            
							var values2 = range2.getValues();
                            values2[0][sheet.getLastColumn()] = "Modification Date";
                            
                        
                             // date of the google sheet creation 
                            /* var d =new Date();
                             var formattedDate = Utilities.formatDate(file.getLastUpdated() , "CET", "yyyy.MM.dd HH:mm:ss.SSS");
                            
                        
                            for (var i = 1; i < values2.length; i++) {  values2[i][values2[i].length - 1] =formattedDate;}
                            range2.setValues(values2);*/
                             
							if (rule.filenameFrom && rule.filenameTo && rule.filenameFrom == file.getName()) {
								var newFilename = Utilities.formatDate(messageDate, config.timezone, rule.filenameTo.replace('%s', message.getSubject()));
								//Logger.log("INFO:           Renaming matched file '" + file.getName() + "' -> '" + newFilename + "'");
								file.setName(newFilename);
							} else if (rule.filenameTo) {
								var newFilename = Utilities.formatDate(messageDate, config.timezone, rule.filenameTo.replace('%s', message.getSubject()));
								//Logger.log("INFO:           Renaming '" + file.getName() + "' -> '" + newFilename + "'");
								file.setName(newFilename);
							}
							file.setDescription("Mail title: " + message.getSubject() + "\nMail date: " + message.getDate() + "\nMail link: https://mail.google.com/mail/u/0/#inbox/" + message.getId());
							Utilities.sleep(config.sleepTime);
							
                            //delete xls or xslx from google Drive -----Keep google spreadsheet files
							if (file.getMimeType() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || "application/vnd.ms-excel") {
								file.setTrashed(true);
							}
						} catch (e) {
							Logger.log(e.lineNumber + ' ' + e);
						}
					}
				}
			}
			thread.addLabel(label);
			if (doArchive) {
				//Logger.log("INFO:     Archiving thread '" + thread.getFirstMessageSubject() + "' ...");
				thread.moveToArchive();
			}
		}
	}
	 end = new Date();
	 var runTime = (end.getTime() - start.getTime()) / 1000;
	//Logger.log("INFO: Finished mail attachment processing after " + runTime + "s");
    
    
}
