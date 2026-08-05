Place the userroles.csv file in a sharepoint location and sync it through Onedrive so all people who are going to use it have access to the source of roles.  Or if its just you  store it locally and update the location.
Line 524 is currently set to read the persons userprofile "C:\users"Username"   feel free to change that to where you want but it requires the csv to work


Update Global config Line 13

# ============================================================
# GLOBAL CONFIG
# ============================================================
$TenantDomain          = "#Domain#"
$DefaultUsageLocation = "#Country"
$DefaultCompany       = "#Companyname#"

Go to line 476 and add in your office locations. Add or remove amount of offices based on your company 


# ============================================================
# OFFICE / LOCATION MAP

#To add a new location copy another one and change the details.
   #     "New Site name" = @{
   #     StreetAddress = "123 High Street"
   #    City          = "Christchurch"
   #    State         = "Canterbury?"
   #     Country       = "New Zealand"
   #  }

# ============================================================
$OfficeMap = @{
    "Site 1" = @{
        StreetAddress = "Street Address 1"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
    "Site 2" = @{
        StreetAddress = "Street Address 2"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
    "Site 3" = @{
        StreetAddress = "Street Address 3"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
    "Site 4" = @{
        StreetAddress = "Street Address 4"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
    "Site 5" = @{
        StreetAddress = "Street Address 5"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
}



To add more locations simply copy  between the " "  below  and name it accordingly.
"
   "Site 2" = @{
        StreetAddress = "Street Address 2"
        City          = "City"
        State         = "State"
        Country       = "Country"
    }
"


Update line 1338   with your company name to add title at top of app
Update #Companyname# with your company name
# ============================================================
# MAIN FORM
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "#Companyname# - Entra / Microsoft 365 User Management"




