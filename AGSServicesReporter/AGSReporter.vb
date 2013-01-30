Imports System.Windows.Forms
Imports System.IO
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.GISClient
Imports ESRI.ArcGIS.Server
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ADF
Imports ESRI.ArcGIS.ADF.Connection


Public Class AGSReporter
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()
        '
        ' TODO: enabled Button
        '
    End Sub

    Protected Overrides Sub OnClick()
        '
        '  TODO: get The AGS Server Details and AGS Services Details
        '
        Dim type As System.Type = System.Type.GetTypeFromCLSID(GetType(ESRI.ArcGIS.Framework.AppRefClass).GUID)
        Dim pGxApp As IGxApplication = TryCast(Activator.CreateInstance(type), ESRI.ArcGIS.CatalogUI.IGxApplication)
        Dim pGxObject As IGxObject = pGxApp.SelectedObject

        'test Connection And Connect (If Required)
        If InStr(pGxObject.Category, "ArcGIS Server") > 0 Then

            'get File Browser Dialog
            Dim folderBrowserDialog As New FolderBrowserDialog
            With folderBrowserDialog
                .Description = "select Folder"
                .ShowNewFolderButton = True
                If .ShowDialog = DialogResult.OK Then
                    Dim folder As String = folderBrowserDialog.SelectedPath
                    deleteCSVFiles(folder)
                Else
                    System.Windows.Forms.MessageBox.Show("no Folder Selected", "error", MessageBoxButtons.OKCancel)
                    Exit Sub
                End If
            End With

            'get AGS Server Details
            Dim hostName As String = Trim(Replace(pGxObject.Category, "ArcGIS Server", ""))
            'System.Windows.Forms.MessageBox.Show(hostName, "Status", MessageBoxButtons.OK)
            Dim pGxAGSConnection As IGxAGSConnection = pGxObject

            'if not Connected            
            If Not pGxAGSConnection.IsConnected Then
                pGxAGSConnection.Connect()
            End If

            'System.Windows.Forms.MessageBox.Show("connected AGS To " + hostName, "Status", MessageBoxButtons.OK)
            Dim pAGSServerConnectionName As IAGSServerConnectionName = pGxAGSConnection.AGSServerConnectionName
            Dim pPs As IPropertySet = pAGSServerConnectionName.ConnectionProperties
            Dim connectionFile As String = pPs.GetProperty("connectionfile")

            Dim pAGSServerConnectionFact2 As IAGSServerConnectionFactory2 = New AGSServerConnectionFactory
            Dim pAGSServerConnection As IAGSServerConnection = pAGSServerConnectionFact2.OpenFromFile(connectionFile, 0)

            'check For AGS Admin Connection
            Dim pAGSServerConnectionAdmin As IAGSServerConnectionAdmin = pAGSServerConnection
            Dim pEnumServerObjectName As IAGSEnumServerObjectName = pAGSServerConnection.ServerObjectNames
            Dim pAGSServerObjectName As IAGSServerObjectName = pEnumServerObjectName.Next()

            'loop AGS Map Services
            Do While Not pAGSServerObjectName Is Nothing

                'try Getting The Service Info
                Try

                    'work With The Services
                    Dim pSOM As IServerObjectManager = pAGSServerConnectionAdmin.ServerObjectManager
                    Dim pServerContext As IServerContext = pSOM.CreateServerContext(pAGSServerObjectName.Name, pAGSServerObjectName.Type)
                    'Dim pServerContext As IServerContext = pSOM.CreateServerContext("", "")
                    Dim pSOConfig As IServerObjectConfiguration4 = pAGSServerConnectionAdmin.ServerObjectConfiguration(pAGSServerObjectName.Name, pAGSServerObjectName.Type)
                    Dim pSOConfigInfo As IServerObjectConfigurationInfo = pSOM.GetConfigurationInfo(pAGSServerObjectName.Name, pAGSServerObjectName.Type)
                    Dim pSOProps As IPropertySet = pSOConfig.Properties()

                    'get The Filename
                    Dim FN As String = folderBrowserDialog.SelectedPath + "\" + pAGSServerObjectName.Type + "_" + UCase(hostName) + "_P" + pSOProps.Count.ToString + ".csv"

                    'get Service Properties, Sprt and Populate Header for File
                    Dim header As String = "ServiceName,CleanupTimeout,Description,IdleTimeout,InstancesPerContainer,IsolationLevel,IsPooled,MaxInstances,MinInstances,Name,ServiceKeepAliveInterval,StartupTimeout,StartupType,TypeName,UsageTimeOut,WaitTimeout"
                    Dim nameObj As New Object
                    Dim valueObj As New Object
                    Dim i As Long = 0
                    pSOProps.GetAllProperties(nameObj, valueObj)
                    Dim serviceList As New List(Of String)
                    For i = 0 To pSOProps.Count - 1
                        serviceList.Add(nameObj(i))
                    Next
                    serviceList.Sort()

                    'populate Header With non-Passive Properties
                    For Each srv In serviceList
                        header = header + "," + srv
                    Next

                    'get Service Parameters And Loop
                    Dim details As String = Chr(34) + pAGSServerObjectName.Name.ToString + Chr(34)
                    'get ServerObjectConfig Detaila
                    details = details + "," + Chr(34) + pSOConfig.CleanupTimeout.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.Description.ToString + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.IdleTimeout.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.InstancesPerContainer.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.IsolationLevel.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.IsPooled.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.MaxInstances.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.MinInstances.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.Name.ToString + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.ServiceKeepAliveInterval.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.StartupTimeout.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.StartupType.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.TypeName.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.UsageTimeout.ToString() + Chr(34)
                    details = details + "," + Chr(34) + pSOConfig.WaitTimeout.ToString() + Chr(34)

                    'get ServerObject ProprtySet
                    pSOProps.GetAllProperties(nameObj, valueObj)
                    For Each srv In serviceList
                        details = details + "," + pSOProps.GetProperty(srv).ToString
                    Next

                    'record Details In File
                    recordDetails(FN, header, details)

                    'capture Data Sources
                    If pAGSServerObjectName.Type = "MapServer" Then

                        'get MapServer and Layer Information
                        Dim pMapServer As IMapServer3 = pServerContext.ServerObject
                        Dim pMapServerInfo As IMapServerInfo = pMapServer.GetServerInfo(pMapServer.DefaultMapName)
                        Dim pMapLayerInfos As IMapLayerInfos = pMapServerInfo.MapLayerInfos

                        'get Layer Info of Map Layers
                        Dim mxd As String = folderBrowserDialog.SelectedPath + "\" + pAGSServerObjectName.Type + "_" + Replace(pAGSServerObjectName.Name, "/", "_") + "_DS.csv"
                        getDataFromMxd(mxd, pMapLayerInfos, pMapServer)

                    End If

                    'reFresh The Folder
                    pGxApp.Refresh(FN)

                    'release Server Context
                    pServerContext.ReleaseContext()

                Catch

                    'next Service
                End Try

                'next ServerObject
                pAGSServerObjectName = pEnumServerObjectName.Next()
            Loop

            'disConnect AGS Connection
            pGxAGSConnection.Disconnect()

            'comPlete
            System.Windows.Forms.MessageBox.Show("ags Reporter:" + vbNewLine + "HostName: " + hostName, "Complete", MessageBoxButtons.OK)

        Else
            System.Windows.Forms.MessageBox.Show("not an AGS Connection", "Status", MessageBoxButtons.OK)
        End If

        'reFresh Catalog
        pGxApp.Refresh(0)

    End Sub

    Protected Overrides Sub OnUpdate()
        Enabled = My.ArcCatalog.Application IsNot Nothing
    End Sub


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


    'record Details In File Of AGS Service Type
    Public Sub recordDetails(ByVal fn As String, ByVal header As String, ByVal details As String)
        Dim agsStreamWriter As StreamWriter = New StreamWriter(fn, FileMode.Append)
        Dim fileInfo As FileInfo = New FileInfo(fn)
        If fileInfo.Length = 0 Then
            agsStreamWriter.WriteLine(header)
            agsStreamWriter.WriteLine(details)
        Else
            agsStreamWriter.WriteLine(details)
        End If
        agsStreamWriter.Close()
        agsStreamWriter.Dispose()
    End Sub

    'delete CSV File
    Public Sub deleteCSVFiles(ByVal folder As String)
        Dim dInfo As DirectoryInfo = New DirectoryInfo(folder)
        Dim fArrayInfo As FileInfo() = dInfo.GetFiles("*.csv")
        Dim f As FileInfo
        For Each f In fArrayInfo
            File.Delete(folder + "\" + f.Name)
        Next
    End Sub

    'get Data Sources From MXD
    Public Sub getDataFromMxd(ByVal mxd As String, ByVal pMapLayerInfos As IMapLayerInfos, ByVal pMapServer As IMapServer3)

        'loop Layers
        For n As Integer = 0 To pMapLayerInfos.Count - 1

            'pMapLayerInfo
            Dim pMapLayerInfo As IMapLayerInfo3 = pMapLayerInfos.Element(n)

            'get Some Details For The Layer
            'feature Layer
            Dim header As String = "Server,Instance,Database,Dataset"
            If pMapLayerInfo.IsFeatureLayer Then

                ' Access the source feature class.
                Dim dataAccess As IMapServerDataAccess = DirectCast(pMapServer, IMapServerDataAccess)
                Dim fc As IFeatureClass = DirectCast(dataAccess.GetDataSource(pMapServer.DefaultMapName, n), IFeatureClass)
                Dim pDs As IDataset = fc
                Dim pPs As IPropertySet = pDs.Workspace.ConnectionProperties
                Dim namesObj As Object
                Dim valuesObj As Object
                pPs.GetAllProperties(namesObj, valuesObj)

                'file Or ArcSDE
                If pPs.Count = 1 Then
                    'get Parameters
                    Dim detail As String = pDs.Workspace.PathName + ",,," + pDs.BrowseName
                    'write To File
                    recordDataSources(mxd, header, detail)
                Else
                    'get Parameters
                    Dim server As String = valuesObj(0)
                    Dim instance As String = valuesObj(1)
                    Dim dataBase As String = valuesObj(2)
                    Dim dsName As String = pDs.BrowseName
                    Dim detail As String = server + "," + instance + "," + dataBase + "," + dsName
                    'write To File
                    recordDataSources(mxd, header, detail)
                End If

                'not A Feature Layer
            Else


                'loop Composite Layer
                For i As Integer = 0 To pMapLayerInfo.SubLayers.Count - 1
                    Dim pMapLayerInfo1 As IMapLayerInfo3 = pMapLayerInfos.Element(pMapLayerInfo.SubLayers.Element(i))


                    'if Feature Layer
                    If pMapLayerInfo1.IsFeatureLayer Then

                        ' Access the source feature class.
                        Dim detail As String = pMapLayerInfo1.Type.ToString + ",,," + pMapLayerInfo1.Name.ToString

                        'write To File
                        recordDataSources(mxd, header, detail)

                    End If


                Next

            End If

        Next

    End Sub

    'record Data Source Details
    Public Sub recordDataSources(ByVal fn As String, ByVal header As String, ByVal details As String)
        Dim agsStreamWriter As StreamWriter = New StreamWriter(fn, FileMode.Append)
        Dim fileInfo As FileInfo = New FileInfo(fn)
        If fileInfo.Length = 0 Then
            agsStreamWriter.WriteLine(header)
            agsStreamWriter.WriteLine(details)
        Else
            agsStreamWriter.WriteLine(details)
        End If
        agsStreamWriter.Close()
        agsStreamWriter.Dispose()
    End Sub

End Class
