Option Explicit On 
Option Strict On

Namespace CompuMaster.Data

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     LDAP access to retrieve data
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[baldauf]	2005-07-02	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class Ldap

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Returns different information on all the users matching the filter
        '''     expression within the given domain as contents of a DataTable
        '''
        '''     The Table contains the following columns:
        '''     - UserName      User's accountname
        '''     - FirstName     First name
        '''     - LastName      Last name
        '''     - DiplayName    Diplayed name
        '''     - Title         Position
        '''     - EMail         E-Mail address
        '''     - Phone         Phone number
        '''     - MobilePhone   Cell / mobile phone number
        '''     - VoIPPhone     VoIP phone number
        '''     - Street        Street and house number
        '''     - ZIP           Zip / postal code
        '''     - City          City name
        '''     - Country       Country name
        '''     - Company       Company name
        '''     - Department    Department name
        '''     - Initials      The initials of the user
        '''
        '''     Note that any field except "UserName" is optional.
        '''     All fields are of type String.
        '''     Each user account is represented by a DataRow.
        '''     
        ''' </summary>
        ''' <param name="domain">The domain from which to gather the information</param>
        ''' <param name="SearchFilterExpression">The filter expression for specific selection purposes.
        '''             For valid filter expressions see the documentation about
        '''             System.DirectoryServices.DirectorySearcher.Filter</param>
        ''' <returns>A DataTable containing the information, Nothing if an error occurs during execution</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[abaldauf]	2005-07-02	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function QueryUsers(ByVal domain As String, ByVal SearchFilterExpression As String) As DataTable
            Return CompuMaster.Data.LdapTools.QueryUsers(domain, SearchFilterExpression)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Returns different information on all the users with the given account
        '''     name within the given domain as contents of a DataTable
        '''
        '''     The Table contains the following columns:
        '''     - UserName      User's accountname
        '''     - FirstName     First name
        '''     - LastName      Last name
        '''     - DiplayName    Diplayed name
        '''     - Title         Position
        '''     - EMail         E-Mail address
        '''     - Phone         Phone number
        '''     - MobilePhone   Cell / mobile phone number
        '''     - VoIPPhone     VoIP phone number
        '''     - Street        Street and house number
        '''     - ZIP           Zip / postal code
        '''     - City          City name
        '''     - Country       Country name
        '''     - Company       Company name
        '''     - Department    Department name
        '''     - Initials      The initials of the user
        '''
        '''     Note that any field except "UserName" is optional.
        '''     All fields are of type String.
        '''     Each user account is represented by a DataRow.
        '''     
        ''' </summary>
        ''' <param name="domain">The domain from which to gather the information</param>
        ''' <param name="UserAccountName">The account name for which to search</param>
        ''' <returns>A DataTable containing the information, Nothing if an error occurs during execution</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        '''     [abaldauf]  2005-07-02  Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function QueryUsersByAccountName(ByVal domain As String, ByVal UserAccountName As String) As DataTable
            Return CompuMaster.Data.LdapTools.QueryUsersByAccountName(domain, UserAccountName)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Query the LDAP
        ''' </summary>
        ''' <param name="domain">The domain name which will be used as LDAP server name (to query the domain controller)</param>
        ''' <param name="searchFilterExpression">A search expression to filter the results</param>
        ''' <returns>A datatable containing all data as strings</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[AdminSupport]	07.10.2005	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function Query(ByVal domain As String, ByVal searchFilterExpression As String) As DataTable
            Return CompuMaster.Data.LdapTools.Query(domain, searchFilterExpression)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Returns different information on all the users with the given first
        '''     and / or last name within the given domain as contents of a DataTable
        '''
        '''     The Table contains the following columns:
        '''     - UserName      User's accountname
        '''     - FirstName     First name
        '''     - LastName      Last name
        '''     - DiplayName    Diplayed name
        '''     - Title         Position
        '''     - EMail         E-Mail address
        '''     - Phone         Phone number
        '''     - MobilePhone   Cell / mobile phone number
        '''     - VoIPPhone     VoIP phone number
        '''     - Street        Street and house number
        '''     - ZIP           Zip / postal code
        '''     - City          City name
        '''     - Country       Country name
        '''     - Company       Company name
        '''     - Department    Department name
        '''     - Initials      The initials of the user
        '''
        '''     Note that any field except "UserName" is optional.
        '''     All fields are of type String.
        '''     Each user account is represented by a DataRow.
        '''     
        ''' </summary>
        ''' <param name="domain">The domain from which to gather the information</param>
        ''' <param name="UserFirstName">The first name for which to search (may be empty or nothing if last name is given)</param>
        ''' <param name="UserLastName">The last name for which to search (may be empty or nothing if first name is given)</param>
        ''' <returns>A DataTable containing the information, Nothing if an error occurs during execution</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        '''     [abaldauf]  2005-07-02  Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function QueryUsersByName(ByVal domain As String, ByVal UserFirstName As String, ByVal UserLastName As String) As DataTable
            Return CompuMaster.Data.LdapTools.QueryUsersByName(domain, UserFirstName, UserLastName)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''     Returns different information on all users within the given domain
        '''     as contents of a DataTable
        '''
        '''     The Table contains the following columns:
        '''     - UserName      User's accountname
        '''     - FirstName     First name
        '''     - LastName      Last name
        '''     - DiplayName    Diplayed name
        '''     - Title         Position
        '''     - EMail         E-Mail address
        '''     - Phone         Phone number
        '''     - MobilePhone   Cell / mobile phone number
        '''     - VoIPPhone     VoIP phone number
        '''     - Street        Street and house number
        '''     - ZIP           Zip / postal code
        '''     - City          City name
        '''     - Country       Country name
        '''     - Company       Company name
        '''     - Department    Department name
        '''     - Initials      The initials of the user
        '''
        '''     Note that any field except "UserName" is optional.
        '''     All fields are of type String.
        '''     Each user account is represented by a DataRow.
        '''     
        ''' </summary>
        ''' <param name="domain">The domain from which to gather the information</param>
        ''' <returns>A DataTable containing the information, Nothing if an error occurs during execution</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        '''     [abaldauf]  2005-07-02  Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Shared Function QueryAllUsers(ByVal domain As String) As DataTable
            Return CompuMaster.Data.LdapTools.QueryAllUsers(domain)
        End Function

    End Class

End Namespace
