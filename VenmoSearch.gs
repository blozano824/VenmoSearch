function venmoSearch_v5(){
  var sheet = SpreadsheetApp.openByUrl("ENTER SPREADSHEET URL HERE");
  var max = 50;
  var currentSearch = 0;
  var accountNum = "N/A";
  
  do{
    var searchResults = GmailApp.search("Venmo", currentSearch, max);
    for(var result = 0; result < searchResults.length; result++){
      var thread = searchResults[result];
      var message = thread.getMessages()[0] // get first message
      
      var date = message.getDate()
      var formattedDate = Utilities.formatDate(date, "ET", "MM-dd-YY");
      
      if (message.getFrom() == "Venmo <venmo@venmo.com>"){
        
        var subject = message.getSubject()
        var subjectArray = subject.split(" ") 
        var subjectArrayLength = subjectArray.length
        var transaction = false
        
        
        for (var word = 0; word < subjectArrayLength - 1; word++){
          if (subjectArray[word] == "You" && subjectArray[word + 1] == "paid"){
            var transactionType = "Sent"
            var money = subjectArray[subjectArrayLength - 1];
            money = "-" + money;
            transaction = true;
          }
          
          if (subjectArray[word] == "paid" && subjectArray[word + 1] == "you"){      
            var transactionType = "Received";   
            var money = subjectArray[subjectArrayLength - 1];
            transaction = true;
          }
          
          if (subjectArray[word] == "charge" && subjectArray[word + 1] == "request"){
              
            var money = subjectArray[word - 1]
            if (subjectArray[0] == "You" && subjectArray[1] == "completed"){
              var transactionType = "Charged Sent"
              transaction = true
            }else{
              var transactionType = "Charged Received"
              transaction = true
            }
          }
        }
    
        if (transaction == true){
          var fullName = nameConstructor(transactionType, subjectArray)
          
          var plainBody = message.getPlainBody()
          var plainBodyArray = plainBody.split(" ");
          var plainBodyArrayLength = plainBodyArray.length
          
          for (var index = 0; index < plainBodyArrayLength; index++){
            if (plainBodyArray[index] == "account" && plainBodyArray[index + 1] == "ending" && plainBodyArray[index + 2] == "in"){
              var accountNum = plainBodyArray[index + 3];
              accountNum = accountNum.replace(".","")
              accountNum = accountNum.trim()
            }
          }
          
          plainBodyScanResults = plainBodyScanner(plainBodyArray,plainBodyArrayLength)
          var paymentID = plainBodyScanResults[0]
          var textStart = plainBodyScanResults[1]
          var textEnd = plainBodyScanResults[2]
          var adjusted = plainBodyScanResults[3]
          
          var textString = textFormatter(plainBodyArray, textStart, textEnd, adjusted)
  
          consoleChecker(formattedDate, fullName, money, transactionType, paymentID, accountNum, textString)
          sheet.appendRow([formattedDate, fullName, money, transactionType, paymentID, accountNum, textString]);
        }
      }
    }
    currentSearch += max;
  }while(searchResults.length == max);
}


function consoleChecker(formattedDate, fullName, money, transactionType, paymentID, accountNum, textString){
  Logger.log("Date: " + formattedDate);
  Logger.log("Person: " + fullName);
  Logger.log(transactionType + ": " + money)
  Logger.log("Payment ID: " + paymentID)
  Logger.log("Bank Account Used: " + accountNum)
  Logger.log("Message: " + textString)
}

function nameConstructor(transactionType, subjectArray){
  var nameArray = new Array();
  
  // Finds name of other Venmo user when money was sent
  if (transactionType == "Sent"){
    for (nameCurrent = 2; nameCurrent < subjectArray.length - 1; nameCurrent++){
      nameArray.push(subjectArray[nameCurrent]);
    }
  }
  
  // Finds name of other Venmo user when money was received
  if (transactionType == "Received"){
    for (nameCurrent = 0; nameCurrent < subjectArray.length - 3; nameCurrent++){
      nameArray.push(subjectArray[nameCurrent]);
    }
  }
  
  // Finds name of other Venmo user when money was sent due to a charge
  if (transactionType == "Charged Sent"){
    var nameStart
    var nameEnd
    for (var word = 0; word < subjectArray.length - 1; word++){
      if (subjectArray[word] == "completed"){
        nameStart = word + 1
      }
      if (subjectArray[word] == "charge"){
        nameEnd = word - 1
      }
    }
    for (nameCurrent = nameStart; nameCurrent < nameEnd; nameCurrent++){
      nameArray.push(subjectArray[nameCurrent]);
    }
  }
  
  // Finds name of other Venmo user when money was received due to a charge
  if (transactionType == "Charged Received"){
    for (var word = 0; word < subjectArray.length - 1; word++){
      if (subjectArray[word] == "completed"){
        for (nameCurrent = 0; nameCurrent < word; nameCurrent++){
          nameArray.push(subjectArray[nameCurrent]);
        }
      }
    }
  }
  
  
  var name = nameArray.join(" ")
  name = name.replace("'s","")
  return name;
}

function plainBodyScanner(plainBodyArray,plainBodyArrayLength){
  var ID = "N/A"
  var End = 0;
  var EndAdjusted = 0;
  // Goes through every word in the email from a Venmo Transaction
  for (var index = 0; index < plainBodyArrayLength; index++){
    
    // Finds Payment ID from Venmo Transaction
    if (plainBodyArray[index] == "ID:"){
      var ID = plainBodyArray[index + 1]
    }
    
    // Finds Starting position of text sent in Venmo Transaction
    if (plainBodyArray[index] == "charged" || plainBodyArray[index] == "paid"){
      for(var scan = 1; scan < 6; scan++){
        var str = plainBodyArray[index + scan]
        if (str.search("https://venmo.com") != -1){
          var Start = index + scan + 1;
        }
      }
    }
    
    //Finds Ending position of text sent in Venmo Transaction
    if (plainBodyArray[index] == "Date" && plainBodyArray[index + 1] == "and" && plainBodyArray[index + 2] == "Amount:"){
      End = index - 1;
    }
    if (plainBodyArray[index].search("\nSep") != -1 || plainBodyArray[index].search("\nOct") != -1 || plainBodyArray[index].search("\nNov") != -1 
    || plainBodyArray[index].search("\nDec") != -1 || plainBodyArray[index].search("\nJan") != -1  || plainBodyArray[index].search("\nFeb") != -1
    || plainBodyArray[index].search("\nMar") != -1 || plainBodyArray[index].search("\nApr") != -1  || plainBodyArray[index].search("\nMay") != -1
    || plainBodyArray[index].search("\nJun") != -1 || plainBodyArray[index].search("\nJul") != -1  || plainBodyArray[index].search("\nAug") != -1){
      if (plainBodyArray[index+9].search("\nLike") == 0){
        EndAdjusted = index;
      }
    }
  }
  if (End > 0){
    return [ID, Start, End, 0]
  }else{
    return [ID, Start, EndAdjusted, 1]
  }
}

function textFormatter(plainBodyArray, textStart, textEnd, adjusted){
  var textArray = new Array();
  // Finds text sent in Venmo Transaction using the Starting and Ending points in the email's body array
  if (textStart > 0 && textEnd > 0){
    for(textStart; textStart <= textEnd; textStart++){
      textArray.push(plainBodyArray[textStart]);
    }
  }
  var textString = textArray.join(" ")
  
  // Cleans up text by eliminting extra word Transfer as well as new line characters
  textString = textString.replace("Transfer","")
  if (adjusted == 1){
    textString = textString.slice(0, -3)
  }
  textString = textString.trim()
  return textString
}
