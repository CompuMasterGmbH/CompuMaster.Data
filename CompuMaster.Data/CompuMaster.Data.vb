Option Explicit On 
Option Strict On

Namespace CompuMaster.Data

    ''' <summary>
    '''     CompuMaster common tools and utilities for data exchange
    ''' </summary>
    ''' <remarks>
    ''' PLEASE NOTE: the concept of System.Data.DataTable causes System.OutOfMemoryExceptions at a system-specific limit, depending on installed RAM, RAM usage, RAM fragmentation, size of data in DataTable, etc.
    ''' Some systems with 8 GB RAM installed might be able to handle 8,000,000 rows in a System.Data.DataTable, while other systems might be able to manage more or lesser rows.
    ''' </remarks>
    ''' <copyright>CompuMaster GmbH</copyright>
    Friend Class NamespaceDoc
        'UPDATE FOLLOWING LINE FOR EVERY CHANGE TO TRACK THE VERSION NUMBER INSIDE THIS DOCUMENT
        'Last change on V3.50 - 2009-06-25 JW
    End Class

End Namespace
