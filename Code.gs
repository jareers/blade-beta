/**
 * Copyright Boston Venture Studio LLC https://bvs.net
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


// Key to refer to the minimum number of emails that trigger the add-on's processing
const MIN_EMAILS = "MIN_EMAILS";

// Key to refer to the automatic archiving feature of the add-on
const AUTO_ARCHIVE = "AUTO_ARCHIVE";

// Key to refer to the automatic hourly processing feature of the add-on
const AUTO_PROCESS = "AUTO_PROCESS";

// Key to refer to the label that is applied to emails by the add-on
const LABEL = "LABEL";

// Key to refer to the option for limiting the add-on's operations to the primary inbox only
const ONLY_PRIMARY = "ONLY_PRIMARY";

// Key to refer to the automatic reply feature of the add-on
const AUTO_REPLY = "AUTO_REPLY";

// Key to refer to the text that is sent as an automatic reply by the add-on
const REPLY_TEXT = "REPLY_TEXT";

// Setting default values for the properties.
const defaults = {
  MIN_EMAILS: "0",
  AUTO_ARCHIVE: "true",
  LABEL: "Strangers",
  ONLY_PRIMARY: "false",
  AUTO_REPLY: "false",
  REPLY_TEXT:
    "Your email to [EMAIL_ADDRESS] was removed from inbox and archived by https://blade.net gmail filter.",
  CACHED_CONTACT_LIST: "{}"
};

// Key to store and retrieve a cached contact list in the PropertiesService.
const CACHED_CONTACT_LIST = "CACHED_CONTACT_LIST";

/**
 * This object is used to classify contacts into three categories: 
 * EMAIL_SENT: contacts to whom the user has sent an email.
 * CONTACT: contacts in the user's Google Contacts list.
 * OTHER_CONTACT: other contacts category in Google Contacts list.
**/
const contactTypes = {
  EMAIL_SENT: "EMAIL_SENT",
  CONTACT: "CONTACT",
  OTHER_CONTACT: "OTHER_CONTACT",
};

// Temporary cache for storing user's own email addresses during execution cycle.
let myEmailsCache = null;

function setUserProperty(key, obj) {
  let userPropertiesService = PropertiesService.getUserProperties();
  userPropertiesService.setProperty(key, obj);
}

function setUserProperties(
  minEmails = defaults.MIN_EMAILS,
  autoArchive = defaults.AUTO_ARCHIVE,
  label = defaults.LABEL,
  onlyPrimary = defaults.ONLY_PRIMARY,
  autoReply = defaults.AUTO_REPLY,
  replyText = defaults.REPLY_TEXT
) {

  let saveObject = {
    MIN_EMAILS: minEmails.toString(),
    AUTO_ARCHIVE: autoArchive.toString(),
    LABEL: label,
    ONLY_PRIMARY: onlyPrimary.toString(),
    AUTO_REPLY: autoReply.toString(),
    REPLY_TEXT: replyText,
  };

  let userPropertiesService = PropertiesService.getUserProperties();
  userPropertiesService.setProperties(saveObject);
}

// Function to get a user property by key. If the key doesn't exist, it sets the default value.
function getUserProperty(key) {
  let userProperties = PropertiesService.getUserProperties().getProperties();
  if (
    (userProperties[key] === undefined || userProperties[key] === null) &&
    defaults[key] !== undefined
  ) {
    setUserProperty(key, defaults[key]);
    userProperties = PropertiesService.getUserProperties().getProperties();
  }
  return userProperties[key];
}

// Function to get all user properties and convert them to the correct types
function getUserProperties() {
  let userProperties = {};
  userProperties[MIN_EMAILS] = parseInt(getUserProperty(MIN_EMAILS));
  userProperties[AUTO_ARCHIVE] = getUserProperty(AUTO_ARCHIVE) === "true";
  userProperties[LABEL] = getUserProperty(LABEL);
  userProperties[ONLY_PRIMARY] = getUserProperty(ONLY_PRIMARY) === "true";
  userProperties[AUTO_REPLY] = getUserProperty(AUTO_REPLY) === "true";
  userProperties[REPLY_TEXT] = getUserProperty(REPLY_TEXT);

  return userProperties;
}

/**
 * The createFilter function generates a filter query for the GmailApp search. 
 * It uses the Gmail search syntax to find emails in the user's inbox which 
 * are addressed to the user and are newer than 2 days.
 * 
 * If @param {onlyPrimary} parameter is true, the function will only return 
 * emails from the 'Primary' category of the inbox.
**/
function createFilter(onlyPrimary = true) {
  let inbox_filter = ["to:me", "in:inbox", "newer_than:1d"];

  let filter = inbox_filter.join(" ");
  if (onlyPrimary) {
    filter = filter + " category:primary";
  }
  return filter;
}

/** 
 * processInboxByBlade processes the user's inbox. It finds email threads 
 * matching the created filter in batches of 5.
 * If a batch has less than 5 threads, it stops fetching more. 
 * It then checks each thread for being unsolicited and increments totalMovedCount. 
 * totalMovedCount, the count of potentially moved threads, is returned.
**/
function processInboxByBlade() {
  const maxThreads = 5;
  let startIndex = 0;

  let onlyPrimary = getUserProperty(ONLY_PRIMARY) === "true";
  let searchFilter = createFilter(onlyPrimary);
  let allThreads = [];

  // Make no more than 5 attempts, so as to avoid hitting script 
  // execution time limit (45 secs) or peopleapi contact access 
  // limit per min (35 contacts) put in by Google. 
  // (5 * 5 = 25 threads at max)
  let attempts = 5;
  while (attempts--) {
    let threads = GmailApp.search(searchFilter, startIndex, maxThreads);
    allThreads.push(...threads);
    startIndex = startIndex + maxThreads;
    if (threads.length < maxThreads) {
      break;
    }
  }
  
  Logger.log('Threads# ' + allThreads.length);
  let totalMovedCount = 0;
  allThreads.forEach((element) => {
    totalMovedCount += checkIfUnsolicited(element);
  });

  return totalMovedCount;
}

// Function to check if a given email address belongs to the user
function isMe(fromAddress) {
  let addresses = getEmailAddresses();
  for (i = 0; i < addresses.length; i++) {
    let address = addresses[i],
      r = RegExp(address, "i");
    if (r.test(fromAddress)) {
      return true;
    }
  }
  return false;
}

// Function to get all email addresses of the active user
function getEmailAddresses() {
  if (myEmailsCache === null) {
    let myEmail = Session.getActiveUser().getEmail();
    let myEmails = GmailApp.getAliases();

    myEmails.push(myEmail);
    myEmailsCache = myEmails;
  }
  return myEmailsCache;
}

function getLabelFromName(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

function getCachedContactList() {
  let cachedContactList = getUserProperty(CACHED_CONTACT_LIST);
  return JSON.parse(cachedContactList);
}

function saveCachedContactList(cachedContactList) {
  setUserProperty(CACHED_CONTACT_LIST, JSON.stringify(cachedContactList));
}

function saveContactInCache(senderEmailAddress, contactType) {
  let cachedContactList = getCachedContactList();

  cachedContactList[senderEmailAddress] = contactType;
  saveCachedContactList(cachedContactList);
  return getCachedContactList();
}

// Function to check if a contact exists in the cache
function isContactInCache(senderEmailAddress) {
  let contactTypeInCache = getCachedContactList()[senderEmailAddress];
  return !(contactTypeInCache === null || contactTypeInCache === undefined);
}

/**
 * The isContact function checks if a given email address is a recognized contact for the user. 
 * It checks in the user's Gmail Contacts,and caches them if found, returns boolean accordingly.
**/
function isContact(senderEmailAddress) {
  if (senderEmailAddress === undefined || senderEmailAddress === null) {
    return false;
  }

  let contactType = contactTypes.CONTACT;
  //Initial empty call required by Gmail
  let contact = People.People.searchContacts({
    readMask: "emailAddresses",
  });

  contact = People.People.searchContacts({
    readMask: "emailAddresses",
    query: senderEmailAddress,
  });

  let isContact = contact.results !== undefined && contact.results.length > 0;
  if (isContact) {
    saveContactInCache(senderEmailAddress, contactType);
  }
  return isContact;
}

function extractEmailFromString(inputString) {
  let emailPattern = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
  let match = inputString.match(emailPattern);
  if (match) {
    return match[0];
  } else {
    return null;
  }
}

/** 
 * The checkIfUnsolicited function evaluates whether a particular email thread is unsolicited or not.
 * It checks conditions like only one recipient (the user), sender not in contacts,and have never been replied-to 
 * If the conditions meet, the function labels and optionally archives the thread and sends an auto-reply.
 * If the user has replied to the sender before, the sender is saved in the cache as a sent email contact.
 * The function returns the count of threads moved to the archive.
**/
function checkIfUnsolicited(thread) {
  const userProperties = getUserProperties();
  const messages = GmailApp.getMessagesForThread(thread);
  const recipientCount = messages[0].getTo().split(",").length;
  const label = getLabelFromName(userProperties[LABEL]);
  const senderEmailAddress = extractEmailFromString(messages[0].getFrom());
  const recipientEmailAddress = extractEmailFromString(messages[0].getTo());
  let movedCount = 0;

  if (
    recipientCount == 1 &&
    isMe(recipientEmailAddress) &&
    !(isContactInCache(senderEmailAddress) || isContact(senderEmailAddress))
  ) {
    const myReplyCount = GmailApp.search(
      // Filter emails sent as auto-reply with subject prefix 'Blade.net:'
      "from:me to:" + messages[0].getFrom() + " -subject:Blade.net:" 
    ).length;
    if (myReplyCount == 0) {
      thread.addLabel(label);
      if (userProperties[AUTO_ARCHIVE]) {
        GmailApp.moveThreadToArchive(thread);
        movedCount += 1;
      }
      if (userProperties[AUTO_REPLY]) {
        let emailMessage = userProperties[REPLY_TEXT];
        emailMessage = emailMessage.replace(
          "[EMAIL_ADDRESS]",
          recipientEmailAddress
        );
        GmailApp.sendEmail(
          senderEmailAddress,
          // Do not change the subject prefix "Blade.net:", it is used for filtering auto-reply messages
          "Blade.net: Removed your email from Inbox", 
          emailMessage
        );
      }
    }
    else if (myReplyCount > 0) {
      saveContactInCache(senderEmailAddress, contactTypes.EMAIL_SENT);
    }
  }

  return movedCount;
}

// -------------- Addon UX Code --------------

// Function that generate the Card layout for the addon's homepage
function getHomePageCard(message) {
  const userProperties = getUserProperties();

  let autoArchive = userProperties[AUTO_ARCHIVE];
  let onlyPrimary = userProperties[ONLY_PRIMARY];
  let autoReply = userProperties[AUTO_REPLY];
  let replyText = userProperties[REPLY_TEXT];

  let labelText = CardService.newTextParagraph()
    .setText("Check [Strangers] label for archived emails. People in your contacts will not be processed.");

  let autoArchiveField = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName(AUTO_ARCHIVE)
    .addItem("Auto Archive", AUTO_ARCHIVE, autoArchive);

  let onlyPrimaryField = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName(ONLY_PRIMARY)
    .addItem("Only scan primary Inbox", ONLY_PRIMARY, onlyPrimary);

  let processNowBtn = CardService.newTextButton()
    .setText("Process Now")
    .setOnClickAction(CardService.newAction().setFunctionName("processNow"));

  let fixedFooter =
    CardService.newFixedFooter().setPrimaryButton(processNowBtn);

  let hourlySwitch = CardService.newSwitch()
    .setFieldName(AUTO_PROCESS)
    .setValue(AUTO_PROCESS)
    .setOnChangeAction(
      CardService.newAction().setFunctionName("handleHourlyProcessSwitchChange")
    );

  hourlySwitch.setSelected(isInboxTriggerInstalled());

  let setHourlyProcessSwitch = CardService.newDecoratedText()
    .setTopLabel("Auto-Hourly Process")
    .setText("Setup hourly trigger. Toggling this will save current values.")
    .setWrapText(true)
    .setSwitchControl(hourlySwitch);

  let autoReplyField = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setFieldName(AUTO_REPLY)
    .addItem("Send an auto-reply", AUTO_REPLY, autoReply);

  let replyTextField = CardService.newTextInput()
    .setFieldName(REPLY_TEXT)
    .setTitle("Auto-reply Text")
    .setValue(replyText)
    .setMultiline(true)
    .setHint("Use [EMAIL_ADDRESS] to replace with your email address.");

  let cardHeader = CardService.newCardHeader()
    .setImageUrl(
      "https://cdn.discordapp.com/attachments/1049256202015621160/1129044143612629082/Blade-logo.png"
    )
    .setImageStyle(CardService.ImageStyle.CIRCLE)
    .setTitle("Configure your Blade Addon")
    .setSubtitle("Declutter your Inbox");

  let mainSection = CardService.newCardSection()
    .addWidget(labelText)
    .addWidget(autoArchiveField)
    .addWidget(onlyPrimaryField)
    .addWidget(setHourlyProcessSwitch);

  let autoReplySection = CardService.newCardSection()
    .addWidget(autoReplyField)
    .addWidget(replyTextField)
    .setCollapsible(true)
    .setHeader("Auto-reply Configuration");

  let card = CardService.newCardBuilder()
    .setHeader(cardHeader)
    .setFixedFooter(fixedFooter);

  card.addSection(mainSection);
  card.addSection(autoReplySection);

  return card;
}

function handleHourlyProcessSwitchChange(e) {
  let autoProcess = e.formInput[AUTO_PROCESS];
  let handlerFunc = (autoProcess === AUTO_PROCESS) ? addTrigger: deleteTrigger;
  saveFormValuesAndExecuteHandler(e, handlerFunc);
  return updateHomePageCard();
}

// Function to add the hourly process trigger
function addTrigger(
  minEmails,
  autoArchive,
  label,
  onlyPrimary,
  autoReply,
  replyText
) {
  setUserProperties(
    minEmails,
    autoArchive,
    label,
    onlyPrimary,
    autoReply,
    replyText
  );

  //Install trigger if it doesnt exist
  if (!isInboxTriggerInstalled()) {
    ScriptApp.newTrigger("processInboxByBlade")
      .timeBased()
      .everyHours(1)
      .create();
    Logger.log("Trigger successfully installed");
    return;
  }
  Logger.log("Trigger already installed");
}

// Function to delete the hourly process trigger
function deleteTrigger(
  minEmails,
  autoArchive,
  label,
  onlyPrimary,
  autoReply,
  replyText
) {
  setUserProperties(
    minEmails,
    autoArchive,
    label,
    onlyPrimary,
    autoReply,
    replyText
  );

  const allTriggers = ScriptApp.getProjectTriggers();
  const triggerName = "processInboxByBlade";

  if (allTriggers.length !== 0) {
    for (let i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getHandlerFunction() === triggerName) {
        Logger.log("Deleting trigger at: " + i);
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    }
  }
}

// Function to make sure that card is updated and not pushed as new to avoid unncessary navigation
function updateHomePageCard(message) {
  let card = getHomePageCard(message).build();

  // Return a built ActionResponse that uses the navigation object.
  let nav = CardService.newNavigation().updateCard(card);
  return CardService.newActionResponseBuilder().setNavigation(nav).build();
}

function displayNotification(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(message))
    .build();
}

function isInboxTriggerInstalled() {
  const allTriggers = ScriptApp.getProjectTriggers();
  const triggerName = "processInboxByBlade";

  let isInboxTriggerInstalled = false;
  if (allTriggers.length !== 0) {
    let index = 0;
    while (!isInboxTriggerInstalled && index < allTriggers.length) {
      if (allTriggers[index].getHandlerFunction() === triggerName) {
        isInboxTriggerInstalled = true;
        break;
      }
      index++;
    }
  }
  return isInboxTriggerInstalled;
}

// Function to build the add-on - Entry point for Addon rendering when clicked in Gmail
function buildAddOn() {
  return getHomePageCard().build();
}

function validateLabel(label) {
  return !(label === undefined || label.length <= 0 || label.length >= 30);
}

function validateMinEmails(minEmails) {
  return !(
    minEmails === undefined ||
    isNaN(parseInt(minEmails)) ||
    parseInt(minEmails) === null ||
    parseInt(minEmails) < 0
  );
}

function saveFormValuesAndExecuteHandler(e, handlerFunc) {

  let label = getUserProperty(LABEL);
  let minEmails = getUserProperty(MIN_EMAILS);
  let autoArchive = e.formInput[AUTO_ARCHIVE];
  let onlyPrimary = e.formInput[ONLY_PRIMARY];
  let autoReply = e.formInput[AUTO_REPLY];
  let replyText = e.formInput[REPLY_TEXT];

  if (!validateLabel(label)) {
    label = getUserProperty(LABEL);
  }

  if (!validateMinEmails(minEmails)) {
    minEmails = getUserProperty(MIN_EMAILS);
  }
  handlerFunc(
    minEmails,
    autoArchive === AUTO_ARCHIVE,
    label,
    onlyPrimary === ONLY_PRIMARY,
    autoReply === AUTO_REPLY,
    replyText
  );
  Logger.log('Properties saved');
  return null;
}

// Function to process the inbox immediately upon user request
function processNow(e) {
  saveFormValuesAndExecuteHandler(e, setUserProperties);

  let totalMovedCount = processInboxByBlade();
  if (totalMovedCount > 0) {
    return displayNotification(
      `Inbox processed, ${totalMovedCount} emails identified! Please refresh your Inbox.`
    );
  }

  return displayNotification("No such emails found!");
}