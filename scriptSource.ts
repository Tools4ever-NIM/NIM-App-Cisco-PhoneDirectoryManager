import { nim } from "./nim";


//TODO List
// - Finalize Target Actions
// - Determine best way to store configuration versus hardcoding in script
// - Check Target action error handling
// - Test

// #region Configuration
  const readOnly = false
  const systemname_AD = "AD"
  const systemname_CUCM = "CiscoUCM"
  const systemname_Unity = "CiscoUnity"
  const parkedmailbox_ExtensionUpper = 9000
  const parkedmailbox_ExtensionLower = 9999
  const ldap_enabled = true
  const ldap_enabled_on_transfer = true
  const add_smtp_new_user = true
  const add_unifiedmessaging_user = true
  const UMExternalServiceId = 'ff285289-df56-4404-96f6-3eff87d73637'
// #endregion

// #region Private Functions
  /**
   * Retrieves the current associated devices for the owner specified
   * @param {string} OwnerId - The unique identifier of the owner whose devices are to be fetched.
   * @returns List of associated devices
   */
  async function getOwnerAssociatedDevices(OwnerId: string) {
    return await nim.filterExecute(
      "App_Cisco_Script_GetCUCMUserAssociatedDevices",
      { UserId: OwnerId }
    )
  }

  /**
   * Obtains a list of associated devices for the owner and removes them
   * @param {string} OwnerId - The unique identifier of the owner whose devices are removed
   */
  async function removeOwnerDevices(OwnerId: string) {
    const devices = await getOwnerAssociatedDevices(OwnerId)

    if (devices && devices?.length > 0) {
      nim.logInfo(
        `Owner [${OwnerId}] has ${devices?.length} Devices`
      )

      // Loop over each Devices
      for (const device of devices) {
        nim.logInfo(`Device Name - [${device.name}]`)

        // Skip Physical Phones
        if (!device.name.startsWith("SEP")) {
          nim.logInfo(`Remove Device [${device.name}] - PKID [${device.pkid}]`)
          if(!readOnly) {
            await nim.targetSystemFunctionRun(systemname_CUCM, 'PhonesDelete',{ uuid: device.pkid, name: device.name})
            await nim.targetSystemFunctionRun(systemname_CUCM, 'EndUserDeviceMapsDelete',{ pkid: device.pkid })
          }
        } else {
          nim.logInfo(`Skipping Device Removal [${device.name}] - PKID [${device.pkid}]`)
        }
      }
    } else {
      nim.logInfo(`No devices found for owner [${OwnerId}]`)
    }
  }
  
  /**
   * Retrieves AD User Account for specified sAMAccountName
   * @param {string} sAMAccountName - sAMAccountName of user to return
   * @param {boolean} ignoreError - If True, ignore errors
   * @returns AD User Account
   */
  async function getADUser(sAMAccountName: string, ignoreError?: boolean) {
    const adUser = await nim.filterExecute("App_Cisco_Script_GetADUser", {
      sAMAccountName: sAMAccountName,
    })
    
    if (adUser && adUser.length == 1) {
      nim.logInfo(`Found AD user [${adUser[0].objectGUID}]`)
      return adUser[0]
    } else if (adUser && adUser.length > 1 && !ignoreError) {
      nim.logError(`Found multiple accounts for [${sAMAccountName}]`)
      throw new RangeError(`Found multiple accounts for [${sAMAccountName}]`)
    }

    if(!ignoreError) {
      nim.logError(`Cannot find AD user for [${sAMAccountName}]`)
      throw new ReferenceError(`Cannot find AD user for [${sAMAccountName}]`)
    }

    return null
  }

  /**
   * Retrieves Building Details for specified Building ID
   * @param {number} BuildingID - Building ID to retrieve
   * @returns Building Details
   */
  async function getBuilding(BuildingID: number) {
    const Building = await nim.filterExecute("App_Cisco_Script_GetBuilding", {
      BuildingID: BuildingID,
    })
    
    if (Building && Building.length == 1) {
      nim.logInfo(`Found Building [${Building[0].BuildingID}]`)
      return Building[0]
    } else if (Building && Building.length > 1) {
      nim.logError(`Found multiple buildings for [${BuildingID}]`)
      throw new RangeError(`Found multiple buildings for [${BuildingID}]`)
    }

    return null
  }

  /**
   * Retrieves Phone Templates for specified Building ID
   * @param {number} BuildingID - Building ID to retrieve
   * @returns Phone Templates
   */
  async function getPhoneTemplates(BuildingID: number) {
    const PhoneTemplates = await nim.filterExecute("App_Cisco_Script_GetPhoneTemplates", {
      BuildingID: BuildingID,
    })
    
    if (PhoneTemplates && PhoneTemplates.length > 0) {
      nim.logInfo(`Found Phone Templates for Building [${PhoneTemplates[0].BuildingID}]`)
      return PhoneTemplates
    }

    throw new Error(`Failed to retrieve Phone Templates for Building [${BuildingID}]`)
  }

  /**
   * Retrieves Universal Device Template
   * @param {string} UUID - Universal Device Template to retrieve
   * @returns Universal Device Template details
   */
  async function getUniversalDeviceTemplate(UUID: string) {
    const Templates = await nim.filterExecute("App_Cisco_Script_GetCUCMUniversalTemplate", {
      UUID: UUID,
    })
    
    if (Templates && Templates.length == 1) {
      nim.logInfo(`Found Universal Device Template [${Templates[0].uuid}]`)
      return Templates[0]
    } else if (Templates && Templates.length > 1) {
      nim.logError(`Found multiple universal device templates for [${UUID}]`)
      throw new RangeError(`Found multiple universal device templates for [${UUID}]`)
    }

    throw new Error(`Failed to retrieve universal device template for [${UUID}]`)
  }

  /**
   * Retrieves Product
   * @param {string} Enum - Product Enum to retrieve
   * @returns Product details
   */
  async function getProduct(Enum: string) {
    const Products = await nim.filterExecute("App_Cisco_Script_GetCUCMProduct", {
      Enum: Enum,
    })
    
    if (Products && Products.length == 1) {
      nim.logInfo(`Found Product [${Products[0].enum}]`)
      return Products[0]
    } else if (Products && Products.length > 1) {
      nim.logError(`Found product for [${Enum}]`)
      throw new RangeError(`Found products for [${Enum}]`)
    }

    throw new Error(`Failed to get products for [${Enum}]`)
  }

  /**
     * Retrieves Unity User Account for specified alias
     * @param {string} alias - Alisa used to search for user 
     * @param {boolean} ignoreError - If True, ignore errors
     * @returns Unity User Account
     */
  async function getUnityUser(alias: string, ignoreError: boolean) {
    let unityUser = await nim.filterExecute("App_Cisco_Script_GetUnityUser", {
      Alias: alias
    })

    if (unityUser && unityUser.length == 1) {
      nim.logInfo(`Found Unity user [${unityUser[0].ObjectId}]`)
      return unityUser[0]
    } else if (unityUser && unityUser.length > 1 && !ignoreError) {
      nim.logError(`Found multiple accounts for alias [${alias}]`)
      throw new RangeError(`Found multiple accounts for extension [${alias}]`)
    }

    if(!ignoreError) {
      nim.logError(`Cannot find Unity user for alias [${alias}]`)
      throw new ReferenceError(`Cannot find Unity user for alias [${alias}]`)
    }

    return null
  }

  /**
     * Retrieves Unity User Account for specified extension
     * @param {string} Extension - Extension used to search for user 
     * @returns Unity User Account
     */
  async function getUnityUserByExtension(Extension: string) {
    const unitUser = await nim.filterExecute("App_Cisco_Script_GetUnityUserByExtension", {
      DtmfAccessId: Extension
    })

    if (unitUser && unitUser.length == 1) {
      nim.logInfo(`Found Unity user [${unitUser[0].ObjectId}]`)
      return unitUser[0]
    } else if (unitUser && unitUser.length > 1) {
      nim.logError(`Found multiple accounts for extension [${Extension}]`)
      throw new RangeError(`Found multiple accounts for extension [${Extension}]`)
    }

    nim.logError(`Cannot find Unity user for extension [${Extension}]`)
    throw new ReferenceError(`Cannot find Unity user for extension [${Extension}]`)
  }

  /**
     * Checks to see if extension is already being assigned
     * @param {string} Extension - Extension used to search for user 
     * @returns {boolean} - True, if assigned
     */
  async function checkExtensionAssigned(Extension: string) {
    const unitUser = await nim.filterExecute("App_Cisco_Script_GetUnityUserByExtension", {
      DtmfAccessId: Extension
    })

    return !!(unitUser && unitUser.length > 0)
  }

  /**
   * Retrieves specified Phone Line
   * @param {string} UUID - The unique identifier of phone line
   * @returns Phone Line
   */
  async function getCUCMPhoneLine(UUID: string) {
    const PhoneLine = await nim.filterExecute(
      "App_Cisco_Script_GetCUCMPhoneLine",
      { UUID: UUID }
    )

    if (PhoneLine && PhoneLine.length == 1) {
      nim.logInfo(`Found CUCM Line [${PhoneLine[0].uuid}]`)
      return PhoneLine[0]
    } else if (PhoneLine && PhoneLine.length > 1) {
      nim.logError(`Found multiple CUCM Lines for [${UUID}]`)
      throw new RangeError(`Found multiple CUCM Phone Lines for [${UUID}]`)
    }

    nim.logError(`Cannot find CUCM Line for [${UUID}]`)
    throw new ReferenceError(`Cannot find CUCM Line for [${UUID}]`)
  }

  /**
   * Retrieves specified Phone 
   * @param {string} UUID - The unique identifier of phone
   * @returns Phone
   */
  async function getCUCMPhone(UUID: string) {
    const phone = await nim.filterExecute("App_Cisco_Script_GetCUCMPhone", {
      UUID: UUID,
    })

    if (phone && phone.length == 1) {
      nim.logInfo(`Found CUCM Phone [${phone[0].uuid}]`)
      return phone[0]
    } else if (phone && phone.length > 1) {
      nim.logError(`Found multiple CUCM Phone for [${UUID}]`)
      throw new RangeError(`Found multiple CUCM Phone for [${UUID}]`)
    }
    nim.logError(`Cannot find CUCM Phone for [${UUID}]`)
    throw new ReferenceError(`Cannot find CUCM Phone for [${UUID}]`)
  }

  /**
 * Generates a unique random number within a specified range that's not already in a given array.
 * 
 * @param {number} upper - The upper bound of the random number range.
 * @param {number} lower - The lower bound of the random number range.
 * @param {string[]} existingArray - An array of numbers to check against for uniqueness.
 * @return {string} A unique random number not in the existing array.
 */
  async function generateUniqueRandom(upper: number, lower: number, existingArray: string[]) {
    let uniqueRandom
  
    do {
      // Generate a random number between lower and upper (inclusive) and then convert it to a string
      uniqueRandom = Math.floor(Math.random() * (upper - lower + 1) + lower).toString()
    } while (existingArray.includes(uniqueRandom)) // Check if the generated number as string is in the array
    
    return uniqueRandom // Return the unique random number as a string
  }

/**
 * Gets current datetime in a specific string format
 * 
 * @return {string} returns datetime string in YYYY-MM-DD HH:MM:SS:SSS
 */
  async function getCurrentTimestamp() {
    // Get the current date and time
    const now = new Date()

    // Format each part of the date and time
    const year = now.getFullYear()
    const month = (now.getMonth() + 1).toString().padStart(2, '0') // +1 because months are 0-indexed
    const day = now.getDate().toString().padStart(2, '0')
    const hours = now.getHours().toString().padStart(2, '0')
    const minutes = now.getMinutes().toString().padStart(2, '0')
    const seconds = now.getSeconds().toString().padStart(2, '0')
    const milliseconds = now.getMilliseconds().toString().padStart(3, '0')

    // Concatenate everything into the final formatted string
    const formattedDate = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}.${milliseconds}`
    return formattedDate
  }

  /**
 * Formats string to delete anything after first match
 * 
 * @param {string} text - Source text
 * @param {string} pattern - Pattern to match
 * @return {string} Modified string
 */
  function deleteAfterFirstMatch(text: string, pattern: string): string {
    const index = text.indexOf(pattern) // Find the index of the first occurrence of the pattern
    if (index >= 0) {
      // If the pattern is found, return the substring up to that index
      return text.substring(0, index)
    }
    // If the pattern is not found, return the original string
    return text
  }

// #endregion

// #region NIM Functions
  /**
     * Updates the Phone Line assignment for user, additionally removing the current owner
     * @param {string} PhoneLineUUID - The unique identifier of phone line
     * @param {string} PhoneUUID - The unique identifier of phone
     * @param {number} BuildingID - BuildingID   
     * @param {string} ExternalPhoneNumberMask - THe external phone number mask for building
     * @param {string} CurrentUserId - Current owner username of the phone line
     * @param {string} NewUserId - New owner username for the phone line
     * @param {string} NewPhoneLabel - New phone label
     * @param {string} NewPhoneName - New phone name
     */
  export async function UpdateLineAssignment(
    PhoneLineUUID: string,
    PhoneUUID: string,
    BuildingID: number,
    ExternalPhoneNumberMask: string,
    //ProvisionSoftPhone: boolean,
    CurrentUserId: string = '',
    NewUserId: string,
    NewPhoneLabel: string,
    NewPhoneName: string
  ) {
    // #region Validation of Resources
      nim.logInfo("Validating resources prior to executing changes")

      // #region Retrieve Current Owner AD Account
      nim.logInfo("Retrieving current AD user account")
        const currentOwnerADUser = await getADUser(CurrentUserId,true)
      // #endregion

      // #region Retrieve New Owner AD Account
        nim.logInfo("Retrieving new owner AD user account")
        const newOwnerADUser = await getADUser(NewUserId)
      // #endregion

      // #region Retrieve CUCM Phone Line
        nim.logInfo("Retrieving CUCM Phone Line")
        const cucmPhoneLine = await getCUCMPhoneLine(PhoneLineUUID)
      // #endregion

      // #region Retrieve CUCM Phone
        nim.logInfo("Retrieving CUCM Phone")
        const cucmPhone = await getCUCMPhone(PhoneUUID)
      // #endregion

      // #region Retrieve Building
        nim.logInfo("Retrieving Building")
        const Building = await getBuilding(BuildingID)
      // #endregion

      // #region Retrieve Phone Templates
        nim.logInfo("Retrieving Phone Templates")
        const PhoneTemplates = await getPhoneTemplates(BuildingID)
      // #endregion
      
      // #region Get Parked Mailboxes
        const parkedMailboxes = await nim.filterExecute(
          "App_Cisco_Script_GetParkedMailboxes"
        )

        let parkedExtensions = parkedMailboxes.map(obj => obj.UnityUserExtension)
      // #endregion

      nim.logInfo("Validation completed")
    // #endregion
	
    
    // #region Previous Owner Devices
      if(CurrentUserId.length > 0) {    
          nim.logInfo("Check if previous owner has associated devices and remove them")
          await removeOwnerDevices(CurrentUserId)
      }
    // #endregion

    // #region New Owner User Account
      nim.logInfo("Check if new owner has CUCM User Account")
      let newOwnerCUCMUser = {
        pkid: ''
      }
      const newOwnerCUCMUserResults = await nim.filterExecute(
        "App_Cisco_Script_GetCUCMUser",
        { UserId: NewUserId })
      

      if (newOwnerCUCMUserResults && newOwnerCUCMUserResults?.length < 1) {
        nim.logInfo(`New Owner doesn't exist in CUCM, creating user`)
        

        if(!readOnly) {
          let createCUCMUser = await nim.targetSystemFunctionRun(systemname_CUCM, 'EndUsersCreate',{ firstname: newOwnerADUser?.givenName ?? '', lastname: newOwnerADUser?.sn ?? '', userid: newOwnerADUser?.sAMAccountName ?? '' })
          newOwnerCUCMUser.pkid = createCUCMUser.pkid
        }
      } else {
        nim.logInfo(`New Owner exists in CUCM, skipping creating user`)
        newOwnerCUCMUser.pkid = newOwnerCUCMUserResults[0].pkid
      }
    // #endregion

    // #region Update Directory Number Description
      nim.logInfo("Updating Directory Number Description")
      nim.logInfo(
        `UUID: [${cucmPhoneLine.dirn_uuid}] - newOwnerUsername: [${NewUserId}] - newAlertingName: [${NewPhoneName} - description: [${NewPhoneLabel}]`
      )
      
      if(!readOnly) {
        await nim.targetSystemFunctionRun(systemname_CUCM, 'LinesUpdate',{ uuid: cucmPhoneLine.dirn_uuid, description: NewPhoneLabel, alertingName: NewPhoneName, asciiAlertingName: NewPhoneName})
      }
    // #endregion

    // #region Update Device-To-Line Description
      nim.logInfo("Updating Device-To-Line Description")
      nim.logInfo(
        `LineUUID: [${cucmPhoneLine.uuid}] - PhoneUUID: [${cucmPhone.uuid}] - lineIndex: [${cucmPhoneLine.index}] - dirnPattern: [${cucmPhoneLine.dirn_pattern}] - dirnRoutePartitionName: [${cucmPhoneLine.dirn_routePartitionName_text}] - newDescription: [${NewPhoneLabel}] - newName: [${NewPhoneName}] - newExternalCallingMask: [${ExternalPhoneNumberMask}]`
      )
      
      if(!readOnly) {
        await nim.targetSystemFunctionRun(systemname_CUCM, 'PhoneLinesUpdate',{ 
                                                                                                      uuid: cucmPhoneLine.uuid, 
                                                                                                      phone_uuid: cucmPhone.uuid,
                                                                                                      index: cucmPhoneLine.index, 
                                                                                                      dirn_pattern: cucmPhoneLine.dirn_pattern, 
                                                                                                      dirn_routePartitionName_text: cucmPhoneLine.dirn_routePartitionName_text, 
                                                                                                      label: NewPhoneLabel,
                                                                                                      display: NewPhoneName, 
                                                                                                      e164Mask: ExternalPhoneNumberMask })
      }
    // #endregion

    // #region Update the Phone Owner
    nim.logInfo("Updating Phone Owner")
    nim.logInfo(`PhoneUUID: [${cucmPhone.uuid}] - newOwnerUsername: [${NewUserId}]`)
    if(!readOnly) {
      await nim.targetSystemFunctionRun(systemname_CUCM, 'PhonesUpdate',{ uuid: cucmPhone.uuid,ownerUserName_text: NewUserId,removeAllUsersForDevice: 'True' })
    }
    // #endregion

    // #region New Owner Devices
      nim.logInfo("Check if new owner has associated devices and remove")
      await removeOwnerDevices(NewUserId)
    // #endregion

    // #region Update Soft Phone for new user

    // Support for softphones not needed currently, development started but not tested/finished
    /*if(ProvisionSoftPhone) {
      nim.logInfo("Updating Soft Phone for new owner")
      
      if(PhoneTemplates) {
        for (const template of PhoneTemplates) {
          nim.logInfo(`Processing Phone Template [${template.ID}]`)
          let universalDeviceTemplate = await getUniversalDeviceTemplate(template.UniversalDeviceTemplateUuid)
          let product = await getProduct(template.ProductEnum)

          let phoneName = product.devicenameformat
          phoneName = deleteAfterFirstMatch(phoneName, "[")
          phoneName = phoneName.replace('[', '')
          phoneName = phoneName + newOwnerADUser?.sAMAccountName
          phoneName = phoneName.toUpperCase()
          phoneName = phoneName.substring(0, 15)

          nim.logInfo(`Phone Template [${template.ID}] - Phone Name [${phoneName}]`)

          let phoneDescription = universalDeviceTemplate.deviceDescription
          phoneDescription = phoneDescription?.replace('#LN#', newOwnerADUser?.sn ?? '')
          phoneDescription = phoneDescription?.replace('#FN#', newOwnerADUser?.givenName ?? '')
          phoneDescription = phoneDescription?.replace('#ID#', newOwnerADUser?.sAMAccountName ?? '')
          phoneDescription = phoneDescription?.replace('#NAME#', newOwnerADUser?.displayName ?? '')
          phoneDescription = phoneDescription?.replace('#EMAIL#', newOwnerADUser?.mail ?? '')
          phoneDescription = phoneDescription?.replace('#DEPT#', newOwnerADUser?.department ?? '')
          phoneDescription = phoneDescription?.replace('#DIRN#', cucmPhoneLine.dirn_pattern)
          phoneDescription = phoneDescription?.replace('#PRODUCT#', product.name)

          nim.logInfo(`Phone Template [${template.ID}] - Phone Description [${phoneDescription}]`)
          
          if(!readOnly) {
            nim.logInfo(`Creating Phone [${phoneName}]`)  
            let createPhone = await nim.targetSystemFunctionRun(systemname_CUCM, 'PhonesCreate',
              { name: phoneName,
                description: phoneDescription, 
                product: product.enum, 
                class: 'Phone', 
                protocol: 'SIP', 
                protocolSide: 'User',
                useTrustedRelayPoint: universalDeviceTemplate.useTrustedRelayPoint, 
                builtInBridgeStatus: universalDeviceTemplate.builtInBridge, 
                packetCaptureMode: universalDeviceTemplate.packetCaptureMode, 
                certificateOperation: universalDeviceTemplate.certificateOperation,
                deviceMobilityMode: universalDeviceTemplate.deviceMobilityMode, 
                networkLocation: 'Use System Default',
                networkLocale: universalDeviceTemplate.networkLocale, 
                enableExtensionMobility: universalDeviceTemplate.enableExtensionMobility, 
                primaryPhoneName: phoneName, 
                networkHoldMohAudioSourceId: universalDeviceTemplate.networkHoldMohAudioSource, 
                userHoldMohAudioSourceId: universalDeviceTemplate.userHoldMohAudioSource, 
                devicePoolName: universalDeviceTemplate.devicePool,
                phoneTemplateName: universalDeviceTemplate.phoneButtonTemplate, 
                callingSearchSpaceName: universalDeviceTemplate.callingSearchSpace, 
                locationName: universalDeviceTemplate.location, 
                mediaResourceListName: universalDeviceTemplate.mediaResourceGroupList, 
                sipProfileName: universalDeviceTemplate.sipProfile,
                ownerUserName_text: NewUserId
              }) 
         
            nim.logInfo(`Updating Phone Line for [${phoneName}]`)  
            
            
            let updatePhoneLine = await nim.targetSystemFunctionRun(systemname_CUCM,'PhoneLinesUpdate',
              {
                  uuid: createPhone.uuid, 
                  index: '1', 
                  dirn_pattern: cucmPhoneLine.dirn_pattern, 
                  dirn_routePartitionName: cucmPhoneLine.dirn_routePartitionName, 
                  label: phoneDescription, 
                  display: phoneName, 
                  displayAscii: phoneName, 
                  e164Mask: Building?.ExternalPhoneNumberMask ?? ''
              } )
            
              nim.logInfo(`Associating User [${NewUserId}] to Phone [${phoneName}]`)  
            
            
            let updatePhoneLineUser = await nim.targetSystemFunctionRun(systemname_CUCM,'PhoneLinesUpdate', {
              uuid: createPhone.uuid, 
              index: '1',
              dirn_pattern: cucmPhoneLine.dirn_pattern, 
              dirn_routePartitionName: cucmPhoneLine.dirn_routePartitionName, 
              userId: newOwnerCUCMUser.pkid
            })
        }
        }
      } else {
        nim.logInfo("No phone templates to process")
      }
    
    } else {
      nim.logInfo("Skipping Soft Phone for new owner")
    }*/
    // #endregion

    // #region Update New Owner associated devices with all phone names
      nim.logInfo("Update New Owner Associated devices phone names  (remove all existing mapped users)")
      const newOwnerDevices = await getOwnerAssociatedDevices(NewUserId);
        nim.logInfo(`fkdevice: [${cucmPhoneLine.device_pkid}] - fkenduser: [${newOwnerCUCMUser.pkid}] - tkuserassociation: [1] - RemoveAllUsersForDevice: [True]`)
      if(!readOnly) {
        await nim.targetSystemFunctionRun(systemname_CUCM, 'EndUserDeviceMapsCreate',{ fkdevice: cucmPhoneLine.device_pkid, fkenduser: newOwnerCUCMUser.pkid, tkuserassociation: '1', removeAllUsersForDevice: 'True' })
      }
      
    // #endregion

    // #region Update New Owner Primary Extension
      nim.logInfo("Updating New Owner primary extension")
      nim.logInfo(
        `newOwnerUsername: [${NewUserId}] - dirnPattern: [${cucmPhoneLine.dirn_pattern}] - dirnRoutePartitionName: [${cucmPhoneLine.dirn_routePartitionName}]`
      )
      if(!readOnly) {
        await nim.targetSystemFunctionRun(systemname_CUCM,'PhonesUpdate',{ uuid: cucmPhone.uuid, ownerUserName:NewUserId})
      }
    // #endregion

    // #region Update IPPhone & telephoneNumber for AD User
      nim.logInfo(
        `Updating [ipPhone] and [telephoneNumber] for New Owner to [${cucmPhoneLine.dirn_pattern}]`
      )

      if(!readOnly) {
        await nim.targetSystemFunctionRun(systemname_AD, 'UserUpdate',{ objectGUID: newOwnerADUser?.objectGUID ?? '', ipPhone: cucmPhoneLine.dirn_pattern, telephoneNumber: cucmPhoneLine.dirn_pattern } )
      }
    // #endregion

    // #region Reassign current extension owner, Update AD User
    	let SkipExtensionOwner = false  
    	if(CurrentUserId.length > 0) {   
          nim.logInfo("Checking target extension is taken in Unity")

          if(await checkExtensionAssigned(cucmPhoneLine.dirn_pattern)) {
            let CurrentUnityUser = await getUnityUserByExtension(cucmPhoneLine.dirn_pattern)

            if(CurrentUnityUser.Alias.toLowerCase() !== newOwnerADUser?.sAMAccountName.toLowerCase() && CurrentUnityUser.Alias.length > 0) {

              if(!readOnly) {
                //Get unique advailable parked extension
                nim.logInfo("Generating random parked mailbox extension")
                let uniqueParkedExtension = await generateUniqueRandom(parkedmailbox_ExtensionUpper, parkedmailbox_ExtensionLower, parkedExtensions)
                let currentTimestamp = await getCurrentTimestamp()

                nim.logInfo(`Updating parked mailbox Unity user [${CurrentUnityUser.ObjectId}] to extension [${uniqueParkedExtension}]`)
                let i = 0
                while(true) {
                  try {
                    await nim.targetSystemFunctionRun(systemname_Unity,'userUpdate', {
                      ObjectId: CurrentUnityUser.ObjectId,
                      DtmfAccessId: uniqueParkedExtension
                    } );
                    break;
                  } catch(e) {
                    i++;
                    if(i < 10) {
                      parkedExtensions.push(uniqueParkedExtension)
                      uniqueParkedExtension = await generateUniqueRandom(parkedmailbox_ExtensionUpper, parkedmailbox_ExtensionLower, parkedExtensions)
                    } else {
                      throw new Error("Failed to find unique parked extension after 10 attempts")
                    }
                  }
                }

                nim.logInfo("Storing parked mailbox internally")
                await nim.targetSystemFunctionRun('internal', 'Cisco_MailboxParking_create', { UnityUserObjectId: CurrentUnityUser.ObjectId, UnityUserAlias: CurrentUnityUser.Alias, UnityUserExtension: uniqueParkedExtension, DateCreated: currentTimestamp, Deleted: '0'})


                if(currentOwnerADUser && currentOwnerADUser.sAMAccountName.length > 0) {
                  nim.logInfo("Updating Current Owner AD User Account")
                  nim.logInfo(`objectGUID: [${currentOwnerADUser.objectGUID}] - ipPhone: [${uniqueParkedExtension}] - telephoneNumber: [${uniqueParkedExtension}]`)
                  await nim.targetSystemFunctionRun(systemname_AD, 'UserUpdate',{ objectGUID: currentOwnerADUser?.objectGUID ?? '', ipPhone: uniqueParkedExtension, telephoneNumber: uniqueParkedExtension } )
                }
              }
            } else {
              SkipExtensionOwner = true
            }
          }
        } else { SkipExtensionOwner = true }
    // #endregion

    // #region Check New Owner in Unity
    nim.logInfo("Checking if new owner has unity user account")
    let newOwnerUnityUser = await getUnityUser(NewUserId,true)

    if(newOwnerUnityUser && newOwnerUnityUser.ObjectId.length < 1) {
      nim.logInfo("Creating Unity user account for new owner") 
      if(!readOnly) {
        let LdapType = ldap_enabled ? '3' : '0'

        let newUnityUser = await nim.targetSystemFunctionRun(systemname_Unity, 'userCreate',{ 
            Alias: NewUserId, 
            EmailAddress: newOwnerADUser?.mail ?? '', 
            FirstName: newOwnerADUser?.givenName ?? '', 
            LastName: newOwnerADUser?.sn ?? '', 
            LdapType: LdapType,
            DtmfAccessId: cucmPhoneLine.dirn_pattern, 
            TemplateAlias: Building?.UnityUserTemplateName ?? '', 
            CreateSmtpProxyFromCorp: 'true'
          } )
      
        if((newOwnerADUser?.mail ?? '').length > 0 ) {
          if(add_smtp_new_user) {
            try { 
              await nim.targetSystemFunctionRun(systemname_Unity,'SmtpproxyaddressesCreate',{ SmtpAddress: newOwnerADUser?.mail ?? '', ObjectGlobalUserObjectId: newUnityUser.ObjectId })
            } catch(e) {
              nim.logWarning(`Updating Unity user smtp address failed: ${e}`)
            }
          }

          if(add_unifiedmessaging_user) {
            try {
              await nim.targetSystemFunctionRun(systemname_Unity, 'usersexternalserviceaccountsCreate', {
                ExternalServiceObjectId: UMExternalServiceId,
                EnableCalendarCapability: 'true',
                LoginType: '0',
                EnableMailboxSynchCapability: "true",
                EmailAddressUseCorp: 'true',
                SubscriberObjectId: newUnityUser?.ObjectId ?? ''
              })
            } catch(e) {
              nim.logWarning(`Updating Unity user external service account failed: ${e}`)
            }
          }
        }

        newOwnerUnityUser = await getUnityUser(NewUserId,false)
      }
    } else {
      nim.logInfo("Updating Unity user account for new owner")
      if(!readOnly) {
        
        if(!SkipExtensionOwner) {
        nim.logInfo(`Updating Unity user [${newOwnerUnityUser?.ObjectId}] to extension [${cucmPhoneLine.dirn_pattern}]`)
          await nim.targetSystemFunctionRun(systemname_Unity, 'userUpdate', {
            ObjectId: newOwnerUnityUser?.ObjectId,
            DtmfAccessId: cucmPhoneLine.dirn_pattern
          })
        } else {
          nim.logInfo('Owner is already properly assigned, skipping assignment')
        }
          if((newOwnerADUser?.mail ?? '').length > 0 ) {
            if(add_smtp_new_user) {
              try { 
                await nim.targetSystemFunctionRun(systemname_Unity,'smtpproxyaddressesCreate',{ SmtpAddress: newOwnerADUser?.mail ?? '', ObjectGlobalUserObjectId: newOwnerUnityUser?.ObjectId ?? '' })
              } catch(e) {
                nim.logWarning(`Updating Unity user smtp address failed: ${e}`)
              }
            }

            if(add_unifiedmessaging_user) {
              try {
                await nim.targetSystemFunctionRun(systemname_Unity, 'usersexternalserviceaccountsCreate', {
                  ExternalServiceObjectId: UMExternalServiceId,
                  EnableCalendarCapability: 'true',
                  LoginType: '0',
                  EnableMailboxSynchCapability: "true",
                  EmailAddressUseCorp: 'true',
                  SubscriberObjectId: newOwnerUnityUser?.ObjectId ?? ''
                })
              } catch(e) {
                nim.logWarning(`Updating Unity user external service account failed: ${e}`)
              }
            }
        }
      }
      

    }
    // #endregion

    // #region Update New Owner Call Schedule
      nim.logInfo("Updating Unity call schedule for new owner")
      nim.logInfo(`ObjectId [${newOwnerUnityUser?.CallHandlerObjectId}] - ScheduleSetObjectId [${Building?.UnityUserCallScheduleObjectId}]`)
      if(!readOnly) {
        await nim.targetSystemFunctionRun(systemname_Unity,'userscallhandlersUpdate', {
          ScheduleSetObjectId: Building?.UnityUserCallScheduleObjectId ?? '',
          ObjectId: newOwnerUnityUser?.CallHandlerObjectId ?? ''
        })
      }
    // #endregion

    // #region New Owner Transfer Rules
      nim.logInfo("Checking if User Transfer Rules enabled for Building")
      if(Building?.UnityUserTransferRulesEnabled) {
          nim.logInfo("Updating User Transfer Rules")
          
          nim.logInfo(`CallHandlerObjectId [${newOwnerUnityUser?.CallHandlerObjectId}] - Action [${Building?.UnityUserStandardTransferAction}] - Enabled [${Building?.UnityUserStandardTransferEnabled}]`)
          if(!readOnly) {
            await nim.targetSystemFunctionRun(systemname_Unity,'callhandlertransferoptionsUpdate', {
              TransferOptionType: "Standard",
              Action: Building?.UnityUserStandardTransferAction,
              Enabled: Building?.UnityUserStandardTransferEnabled,
              CallHandlerObjectId: newOwnerUnityUser?.CallHandlerObjectId
            })
          }

          nim.logInfo(`CallHandlerObjectId [${newOwnerUnityUser?.CallHandlerObjectId}] - Action [${Building?.UnityUserClosedTransferAction}] - Enabled [${Building?.UnityUserClosedTransferEnabled}]`)
          if(!readOnly) {
            
            await nim.targetSystemFunctionRun(systemname_Unity,'callhandlertransferoptionsUpdate', {
              TransferOptionType: "Off Hours",
              Action: Building?.UnityUserClosedTransferAction,
              Enabled: Building?.UnityUserClosedTransferEnabled,
              CallHandlerObjectId: newOwnerUnityUser?.CallHandlerObjectId
            })
          }


          nim.logInfo(`CallHandlerObjectId [${newOwnerUnityUser?.CallHandlerObjectId}] - Action [${Building?.UnityUserAlternateTransferAction}] - Enabled [${Building?.UnityUserAlternateTransferEnabled}]`)
          if(!readOnly) {
            await nim.targetSystemFunctionRun(systemname_Unity,'callhandlertransferoptionsUpdate', {
              TransferOptionType: "Alternate",
              Action: Building?.UnityUserAlternateTransferAction,
              Enabled: Building?.UnityUserAlternateTransferEnabled,
              CallHandlerObjectId: newOwnerUnityUser?.CallHandlerObjectId
            })
          }
      }
    // #endregion
  }
// #endregion