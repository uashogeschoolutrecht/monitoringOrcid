
import urllib.request
import urllib.parse
import xml.etree.ElementTree as ET
import re
import pandas as pd
from orcidData import *



ns = {'o': 'http://www.orcid.org/ns/orcid' ,
's' : 'http://www.orcid.org/ns/search' ,
'h': 'http://www.orcid.org/ns/history' ,
'p': 'http://www.orcid.org/ns/person' ,
'pd': 'http://www.orcid.org/ns/personal-details' ,
'a': 'http://www.orcid.org/ns/activities' ,
'e': 'http://www.orcid.org/ns/employment' ,
'c': 'http://www.orcid.org/ns/common' }





def getData(firstName, lastName):
    """Retrieve information of ORCID account using the given first and last name."""
    fullOutput = ''

    # The used query to retrieve information.
    query = f'https://pub.orcid.org/v2.1/search?q=family-name:{ urlEncode(lastName) }+AND+given-names:{ urlEncode(firstName)}'
    root = getTree( query )
    hits = root.findall('s:result' , ns )

    count = 0
    for result in hits:
        count += 1

        data = dict()
        orcidId = result.find('c:orcid-identifier/c:path' , ns ).text
        orcidUrl = "https://pub.orcid.org/v3.0/" + orcidId +  "/record"
        xml = getTree( orcidUrl )


        data['lastName'] = getLastName( xml )
        data['firstName'] = getFirstName( xml )
        data['creationDate'] = getCreationDate( xml )
        data['nrWorks'] = getNumberOfWorks( xml )
        aff = getAffiliations( xml )
        for affiliation in aff:
            # Check if the account is associated to the University of Applied Sciences Utrecht or Hogeschool Utrecht
            if(('University of Applied Sciences Utrecht' in affiliation) or ('Hogeschool Utrecht' in affiliation)):
                fullOutput += f"{ lastName },"
                fullOutput += f"{ firstName },"
                fullOutput += f"{ orcidId },"
                fullOutput += f"{ data.get('lastName' , '' ) },"
                fullOutput += f"{ data.get('firstName' , '' ) },"
                fullOutput += f"{ data.get('creationDate' , '' ) },"
                fullOutput += f"{ data.get('nrWorks' , '' ) },"

                if len(aff) > 0:
                    fullOutput += f"{ aff[0][0] },"
                    fullOutput += f"{ aff[0][1] },"
                else:
                    fullOutput += ',,'

                if len(aff) > 1:
                    fullOutput += f"{ aff[1][0] },"
                    fullOutput += f"{ aff[1][1] }\n"
                else:
                    fullOutput += ',\n'
                if count == 3:
                    break

    return fullOutput

# Write the output to a CSV file.
out = open( 'result.csv' , 'w' )
out.write( 'lastName,firstName,orcid,OrcidlastName,OrcidfirstName,creationDate,nrWorks,organisation1,department1,organisation2,department2\n' )

# Open the Excel file which contains the names used to retrieve data.
xl = pd.ExcelFile( 'researchers.xlsx'  )
df = xl.parse( 'Sheet1' )
for index , column in df.iterrows():
    if pd.notnull( column['lastName'] ):
        out.write( getData( urllib.parse.quote( column['firstName'] )  , urllib.parse.quote( column['lastName']  ) ) )

out.close()
