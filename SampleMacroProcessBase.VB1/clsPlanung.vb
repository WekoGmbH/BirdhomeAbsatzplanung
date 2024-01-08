Imports Microsoft.Office.Interop.Excel
Imports Sagede.Core.Tools
Public Class clsPlanung
    Private _Artikelnummer As String
    Public Property Artikelnummer() As String
        Get
            Return _Artikelnummer
        End Get
        Set(ByVal value As String)
            _Artikelnummer = value
        End Set
    End Property
    Private dDatum As Date
    Public Property Datum() As Date
        Get
            Return dDatum
        End Get
        Set(ByVal value As Date)
            dDatum = value
        End Set
    End Property
    Private sDatum As String
    Public Property DatumS() As String
        Get
            Return sDatum
        End Get
        Set(ByVal value As String)
            sDatum = value
        End Set
    End Property
    Private cMenge As Decimal
    Public Property Menge() As Decimal
        Get
            Return cMenge
        End Get
        Set(ByVal value As Decimal)
            cMenge = value
        End Set
    End Property
    Private _Plannummer As Integer
    Public Property Plannummer() As Integer
        Get
            Return _Plannummer
        End Get
        Set(ByVal value As Integer)
            _Plannummer = value
        End Set
    End Property


    Private lPlanung As Int32
    Public Property Planung() As Int32
        Get
            Return lPlanung
        End Get
        Set(ByVal value As Int32)
            lPlanung = value
        End Set
    End Property


    Private lPeriode As Int32
    Public Property Periode() As Int32
        Get
            Return lPeriode
        End Get
        Set(ByVal value As Int32)
            lPeriode = value
        End Set
    End Property


End Class
