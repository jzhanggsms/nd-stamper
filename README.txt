ND Doc Stamper — Deployment Guide
==================================

TRIGGER
  OnDocumentBeforeSave fires before each save.
  First save of an ND doc  -> inserts stamp at end of document.
  Subsequent saves         -> finds stamp via bookmark, updates it.
  Non-ND docs              -> does nothing.

STAMP FORMAT
  ####-####-####. v#.#
  e.g.  1234-5678-9012. v2.1

--------------------------------------------------------------
STEP 1 — Discover your ND property names
--------------------------------------------------------------
Open a live NetDocuments file in Word.
Press Alt+F11, open the Immediate Window (Ctrl+G), paste and run:

  Sub ListNDProperties()
      Dim p As DocumentProperty
      For Each p In ActiveDocument.CustomDocumentProperties
          Debug.Print p.Name & " = " & p.Value
      Next p
  End Sub

Update the ND_PROPS block at the top of commands.js to match.

--------------------------------------------------------------
STEP 2 — Generate a GUID for manifest.xml
--------------------------------------------------------------
  PowerShell:  [guid]::NewGuid()
  Web:         https://guidgenerator.com

--------------------------------------------------------------
STEP 3 — Host the files (HTTPS required)
--------------------------------------------------------------
Replace https://YOUR-SERVER in all files with your actual host.
Options: IIS, Azure Static Web Apps (free tier), SharePoint.

--------------------------------------------------------------
STEP 4 — Deploy via M365 Admin Center
--------------------------------------------------------------
  1. admin.microsoft.com -> Settings -> Integrated apps
  2. Upload custom apps -> select manifest.xml
  3. Assign to all users or a test group
  4. Propagates within ~24h
     (Test immediately via Word: Insert -> Get Add-ins -> My Org)

--------------------------------------------------------------
REQUIREMENTS
--------------------------------------------------------------
  - Word desktop (Microsoft 365, current channel)
  - Word JS API 1.7+
  - OnDocumentBeforeSave does NOT fire in Word Online