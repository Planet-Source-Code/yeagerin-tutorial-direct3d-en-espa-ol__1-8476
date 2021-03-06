VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Begin VB.Form frm_Avion 
   AutoRedraw      =   -1  'True
   Caption         =   "3d Avioncito"
   ClientHeight    =   3915
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6195
   Icon            =   "fly.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer_Render 
      Enabled         =   0   'False
      Interval        =   24
      Left            =   2520
      Top             =   1770
   End
   Begin RMControl7.RMCanvas RMCanvas 
      Height          =   3930
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   6932
   End
   Begin VB.Menu Menu_Anima_Ppal 
      Caption         =   "&Animaci�n"
      Begin VB.Menu mnu_Auto_Anima 
         Caption         =   "&Autom�tica"
         Checked         =   -1  'True
         Shortcut        =   {F5}
      End
      Begin VB.Menu Mnu_Linea 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Anima_Manual 
         Caption         =   "&Manual"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu_Dorno 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Accidente 
         Caption         =   "Accidente"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnu_FPS 
      Caption         =   "FPS"
   End
End
Attribute VB_Name = "frm_Avion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Modelo del avi�n...
Dim M_Avi�n As Direct3DRMFrame3
'Cabina [ para controlar la posici�n el avi�n ]
Dim M_Cabina As Direct3DRMFrame3
'Su animaci�n...
Dim M_AnimaAvi�n As Direct3DRMAnimation2
'Tiempo para la animaci�n...
Dim M_Tiempo As Single
'�Est� volando?
Dim M_Volando As Boolean
Dim VelRot  'Velocidad Rotaci�n
'Aclaraciones:
'======================================================================
'Un frame contiene la geometr�a con posiciones y orientaci�n diferentes
'Un MeshBuilder es pura geometr�a y pueden cambiar la posici�n y orientaci�n
' de sus objetos
'A su vez un frame puede ser vinculado a otro (una escena) y cuando se mueve
' todos los frames se mueven tambi�n junto con los mesh-builders a�adidos a �l.
'Es decir que cuando un frame se mueve todos sus mesh y frames tb lo hacen.
'--=[Es algo as� como el sistema solar]=--------------------------------- :]
'  Donde el sol es el frame ra�z (base),a�adido a �l se encuentran los
'frames childs o hijos que puede contener a la tierra, por ejemplo
' A su vez, a�adido a la tierra (de geometr�a inferior) se encuentran a�adidos
'otros frames hijos como la luna con su propia geometr�a, claro
'imagina que rotamos el Sol; esta rotaci�n generar�a en el frame de la tierra
'rotar alrededor de la �rbita del frame Sol
' Pero podemos rotar la tierra y todo lo que ella contiene dentro del frame del Sol
'==================================================================================
'  RMCanvas.SceneFrame
'Este objeto es la base de todos los dem�s, incluido el ra�z o base root
' es donde generamos nuestro mundo 3D
'---------
'  RMCanvas.CameraFrame
'   Este objeto es un hijo de la escena.
' Como su propio nombre indica determina la posici�n y orientaci�n
'de la c�mara.
' Su valor por defecto es -10 unidades detr�s del eje Z y mirando
' hacia adelante  0,0,0
'---------
'  RMCanvas.DirLightFrame
'  Existen 2 luces por defecto establecidas para ser usadas.
'    * La luz ambiental:
'       + omnidirectional
'       + sin location ni direction
'    * La luz direccional: (hija de la escena-un foco,vamos)
'       + debes usar las funciones setPosition and lookAt para
'           posicionarla y orientarla.
'       + valores por defecto: toward 0,0,0
'---------
'  RMCanvas.DirLight
' Este es el objeto que se establece en el FrameDirLight frame
' Puedes usar el m�todo setColorRGB para cambiar el color de la luz.
'---------
'  RMCanvas.AmbientLight
' Este determina como es la luz ambiental de la escena.
' * setColorRGB cambia el color y la intesidad del light
'   - no uses el blanco porque es "cegador", usa el gris claro :p
'---------
' RMCanvas.Viewport
' Esto describe como funciona la c�mara:
'  * Usa el m�todo setField para determinar como es de cerrada y amplia
' es el area que est�s observando.
'  * Con setFront y setBack determinas desde cu�nta distancia te gustar�a
' ver y c�mo de cerrado observas un objecto.
' Sirve adem�s para obtener objetos que se est�n visualizando, incluso
'se le puede pasar el evento Mouse_Over del rat�n... :]
'---------
'  RMCanvas.Device
'   Esto sirve para declarar el interfaz que controlar� el rendering
' as� como la calidad de dibujo (SetQuality).
' Lo normal es el Gouraund.
'---------
'  RMCanvas.SceneSpeed
'  En unidades por segundo, pueden ser ajustadas = que la rotaci�n
' y la velocidad que afectan a un objecto.Default=30 unidadess/sg.
'---------
'=================================================
'En fin, estas son las propiedades m�s interesantes del control
' RMCanvas paraDirect3D, otras properties te permitir�n por ejemplo
' dibujar sobre la escena como si fueran etiquetas, algo as� como las
' animaciones del CounterStrike para Half-Life :]
'Mirad por ejemplo: DDBackSurface y sus ejemplos.
'=================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not mnu_Anima_Manual.Checked Then Exit Sub
Dim Posici�n_Actual As D3DVECTOR
M_Avi�n.GetPosition M_Cabina, Posici�n_Actual
VelRot = VelRot + 0.0001
Select Case (KeyCode)
    Case 37: 'Izquierda... 37
        'M_Avi�n.AddRotation D3DRMCOMBINE_AFTER, Posici�n_Actual.X + 0.5, Posici�n_Actual.Y + 0.5, Posici�n_Actual.z + 0.5, -VelRot
        'M_Avi�n.AddRotation D3DRMCOMBINE_AFTER, 0#, 15#, 0#, VelRot
        
    Case 38: 'Arriba...
        M_Avi�n.AddRotation D3DRMCOMBINE_AFTER, 0#, 1#, 0#, 0.002
    Case 39: 'Derecha... 39
        M_Avi�n.AddRotation D3DRMCOMBINE_AFTER, 0#, -15#, 0#, VelRot
    Case 40: 'Abajo...
        M_Avi�n.SetPosition M_Cabina, Posici�n_Actual.X, Posici�n_Actual.Y + VelRot, Posici�n_Actual.z
        M_Avi�n.AddRotation D3DRMCOMBINE_AFTER, 0, 1#, 1#, VelRot
    Case 82: 'Rotar
        M_Avi�n.AddRotation D3DRMCOMBINE_AFTER, Posici�n_Actual.X + 0.5, Posici�n_Actual.Y + 0.5, Posici�n_Actual.z + 0.5, VelRot
    Case 76: 'Looping Star!
        M_Avi�n.AddRotation D3DRMCOMBINE_AFTER, Posici�n_Actual.X, Posici�n_Actual.Y, Posici�n_Actual.z + 10, VelRot
    Case Else:
        Exit Sub
End Select
M_AnimaAvi�n.SetFrame RMCanvas.CameraFrame
M_Avi�n.LookAt M_Cabina, Nothing, D3DRMCONSTRAIN_Z
'RMCanvas.CameraFrame.SetPosition Nothing, Posici�n_Actual.x - 10, Posici�n_Actual.y - 10, Posici�n_Actual.z - 10
RMCanvas.CameraFrame.LookAt M_Avi�n, Nothing, D3DRMCONSTRAIN_Z
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    VelRot = 0.002
End Sub

Private Sub Form_Load()
Static Rula As Boolean
'Creamos una variable est�tica para llamadas entre m�dulos
' es decir que no pierda su valor al hacer referencia a �l
' desde otra funci�n y sepamos si rulaba o no el 3D :]
    If Rula = True Then End
   Rula = True
'Mostrar el formulario antes de cargar la escena
    Me.Show
'Hacer los dem�s eventos antes de la escena...
    DoEvents
    VelRot = 0.0002
    Inicio
End Sub
Private Sub Form_Resize()
'Tama�o del Canvas al mismo tama�o que la ventana...
    RMCanvas.width = Me.ScaleWidth
    RMCanvas.height = Me.ScaleHeight
End Sub
Sub Inicio()
Dim Rula As Boolean
Dim sFile As String
'  Vamos a ejecutarlo en modo ventana...
'  encontramos el hardware 3D del 'Display' Primario...
'  si no lo encuentra usar� el Software emulation RGB rasterizer
Rula = RMCanvas.StartWindowed
    If Rula = False Then
        MsgBox "�No puedo iniciar Direct3D con Hardware!" & vbCrLf & _
        vbCrLf & " � Renderizando con software...", vbCritical, "Error de Hardware"
        RMCanvas.Use3DHardware = False
    End If
    Crear_Escena
    Timer_Render.Enabled = True
End Sub


Sub Crear_Escena()
If StrConv(Dir(App.Path & "\land4.x", vbArchive), vbUpperCase) <> "LAND4.X" Then
    MsgBox "No puedo encontrar los mapas de superficie!!", vbCritical, "Error de ficheros"
    End
Else
    Crear_Tierra
    Crear_Modelo
    Crear_Trayectoria_Animaci�n
End If
End Sub

Sub Crear_Tierra()
Dim mb_deTierra As Direct3DRMMeshBuilder3
Static m_tierra As Direct3DRMFrame3
Dim Caja As D3DRMBOX
Dim i As Integer, j As Integer
'- Creamos el frame de la tierra perteneciente al de la escena
    Set m_tierra = RMCanvas.D3DRM.CreateFrame(RMCanvas.SceneFrame)
'- Ahora creamos un objeto para dibujar estructuras...
    Set mb_deTierra = RMCanvas.D3DRM.CreateMeshBuilder()
'- Y cargamos su geometr�a de un ficherofile en el constructor de estructuras
' a�adi�ndolo en el frame.
    mb_deTierra.LoadFromFile App.Path & "\land4.x", 0, 0, Nothing, Nothing
    m_tierra.AddVisual mb_deTierra
'Hacemos que la tierra sea mayor escal�ndola y tomando su extensi�n
' la colocamos en una caja, as� sabremos cu�nto mide exactamente, es como
' tomar las medidas de una fase de un juego para que el personaje
' no se salga de ella. :)
    mb_deTierra.ScaleMesh 10, 8, 10
    mb_deTierra.GetBox Caja
'Metemos sus medidas en la caja, un poco de imaginaci�n :p
Dim RangO As Single
'Ahora tomamos los ejes de la caja...
    RangO = Caja.Max.Y - Caja.Min.Y
'Creamos el color de fondo...
    'RMCanvas.SceneFrame.SetSceneBackground &H6060E0
'La textura de fondo... -by juax- :)
    Dim Textura_Fondo As Direct3DRMTexture3
    Set Textura_Fondo = RMCanvas.CreateUpdateableTexture(256, 256, App.Path & "\cloud3.bmp")
    Textura_Fondo.SetShades Rnd * 50
    Textura_Fondo.GenerateMIPMap
    RMCanvas.SceneFrame.SetSceneBackgroundImage Textura_Fondo
'Y establecemos el color de la luz ambiental por defecto
' del control RMCanvas...
    RMCanvas.AmbientLight.SetColorRGB 0.36, 0.36, 0.36
'Dibujamos las caras seg�n su largo:
    Dim Vertice As D3DVECTOR, Vector_Normal As D3DVECTOR, Y As Single
'Recorremos todos los vertices del MeshBuilder de Tierra
    For i = 0 To mb_deTierra.GetFaceCount() - 1
        Y = Caja.Min.Y
    'Y a su vez por todas las caras...
        For j = 0 To mb_deTierra.GetFace(i).GetVertexCount() - 1
            mb_deTierra.GetFace(i).GetVertex j, Vertice, Vector_Normal
    'Colocamos todos los v�rtices dentro de la caja
            If Vertice.Y > Y Then Y = Vertice.Y
        Next
        If (Y - Caja.Min.Y) / RangO < 0.05 Then
    'Si la cara por la que vamos en el bucle se ve por encima de la tierra, es decir
    ' la c�mara lo ve se dibuja de color potito sino pues de blanco,,,
            Call mb_deTierra.GetFace(i).SetColorRGB((Y - Caja.Min.Y) / RangO, 0.6, 1 - (Y - Caja.Min.Y) / RangO)
        Else
            Call mb_deTierra.GetFace(i).SetColorRGB(0.2 + (Y - Caja.Min.Y) / RangO, 1 - (Y - Caja.Min.Y) / RangO, 0.5)
        End If
    Next
End Sub

Sub Crear_Modelo()
    Dim mb_Avi�n As Direct3DRMMeshBuilder3
'Creamos un MeshBuilder Avi�n para dibujar el modelo...
    Set mb_Avi�n = RMCanvas.D3DRM.CreateMeshBuilder()
'Cargandolo de un fichero.X...
    mb_Avi�n.LoadFromFile App.Path & "\dropship.x", 0, 0, Nothing, Nothing
'Le ajustamos el tama�o...
    mb_Avi�n.ScaleMesh 0.015, 0.008, 0.015
    ' y el color...
    mb_Avi�n.SetColorRGB 0.8, 0.8, 0.8
'Creamos un Frame para representar el modelo...
Set M_Avi�n = RMCanvas.D3DRM.CreateFrame(RMCanvas.SceneFrame)
  ' a�adi�ndolo a la escena
    M_Avi�n.AddVisual mb_Avi�n
    Dim Textura As Direct3DRMTexture3
    Set Textura = RMCanvas.CreateUpdateableTexture(64, 64, App.Path & "\banana.bmp")
    Textura.GenerateMIPMap
    Textura.SetName "Banana"
    M_Avi�n.SetTexture Textura
    M_Avi�n.GetParent.SetTexture Textura
'Lo mismo con su chase (cabina)...
Set M_Cabina = RMCanvas.D3DRM.CreateFrame(RMCanvas.SceneFrame)
'Crea un array de 1000 v�rtices de DirectX � :P ?
'Dim verts(1000) As D3DRMVERTEX
End Sub

Sub Crear_Trayectoria_Animaci�n()
Dim Datos_Trayectoria()
Dim X As Single, Y As Single, z As Single, i As Integer
'Los valores del array son las posiciones del objeto:
    Datos_Trayectoria = Array( _
            -8, 3, -12, _
            -4, 2, -8, _
            -2, 0, -4, _
             9, -1, 7, _
             4, 6, 10, _
            -4, 5, 9, _
             5.5, 3.5, -6.5, _
             2, 5, -10, _
             0, 4, -15, _
            -5, 4, -15, _
            -8, 3, -12)
'...una vez establecida la trayectoria con puntos se puede
' crear la animaci�n:
Set M_AnimaAvi�n = RMCanvas.D3DRM.CreateAnimation()
'Las opciones de la animaci�n est�n establecidas aqu� abajo para que
' se repita continuamente...
    M_AnimaAvi�n.SetOptions D3DRMANIMATION_CLOSED Or D3DRMANIMATION_SPLINEPOSITION Or D3DRMANIMATION_POSITION
Dim Posici�n As D3DRMANIMATIONKEY
'Este bucle va de 10 en 10 cambiando la posici�n del modelo:
    For i = 0 To 10
'Toma los datos escalares del array de posiciones preparado
' especialmente para este "mapa".
        X = Datos_Trayectoria(i * 3)
        Y = Datos_Trayectoria(i * 3 + 1)
        z = Datos_Trayectoria(i * 3 + 2)
'Esto de aqu� abajo se ahorra con: m_AnimaAvi�n.AddPositionKey i, x, y, z
        Posici�n.dvX = X
        Posici�n.dvY = Y
        Posici�n.dvZ = z
        Posici�n.lKeyType = 3
        Posici�n.dvTime = i
'A�adimos el juego de posiciones a la animaci�n para que sepa donde
' tiene que colocar el objeto.
        M_AnimaAvi�n.AddKey Posici�n
    Next
End Sub

Sub Movimiento_Camara_Avi�n(delta As Single)
    Dim Direcci�n As D3DVECTOR
    Dim Direcci�n_Antig�a As D3DVECTOR
    Dim Direcci�n_C�mara As D3DVECTOR
    Dim Dir_Ant_Cam As D3DVECTOR
'Velocidad de la escena...
    RMCanvas.SceneSpeed = 1
'El tiempo de la animaci�n va cambiando con el movimiento de
' la c�mara el delta es el valor de la propia escena...
    M_Tiempo = M_Tiempo + delta
'Colocamos la c�mara seg�n la animaci�n:
    M_AnimaAvi�n.SetFrame RMCanvas.CameraFrame
    M_AnimaAvi�n.SetTime M_Tiempo + 0
'...al igual que el modelo del avi�n...
    M_AnimaAvi�n.SetFrame M_Avi�n
    M_AnimaAvi�n.SetTime M_Tiempo + 0.5
'...y el chase (cabina)...
    M_AnimaAvi�n.SetFrame M_Cabina
    M_AnimaAvi�n.SetTime M_Tiempo + 1
'orientamos la c�mara hacia el avi�n...
    RMCanvas.CameraFrame.LookAt M_Avi�n, Nothing, D3DRMCONSTRAIN_Z
'y el avi�n a su vez lo orientamos hacia la cabina...
    M_Avi�n.LookAt M_Cabina, Nothing, D3DRMCONSTRAIN_Y
'tomamos la orientaci�n de la c�mara...
    RMCanvas.CameraFrame.GetOrientation Nothing, Direcci�n_Antig�a, Dir_Ant_Cam
'y la del avi�n...
    M_Avi�n.GetOrientation Nothing, Direcci�n, Direcci�n_Antig�a
'almacen�ndolas en sus correspondientes variables...
    Direcci�n_Antig�a.X = Direcci�n.X - Direcci�n_Antig�a.X
    Direcci�n_Antig�a.Y = Direcci�n.Y - Direcci�n_Antig�a.Y + 1#
    Direcci�n_Antig�a.z = Direcci�n.z - Direcci�n_Antig�a.z
'Ahora podemos colocar al avi�n (:] eehhee) en su sitio:
    M_Avi�n.SetOrientation Nothing, Direcci�n.X, Direcci�n.Y, Direcci�n.z, Direcci�n_Antig�a.X, Direcci�n_Antig�a.Y, Direcci�n_Antig�a.z
'Fondo..
If mnu_Accidente.Checked Then _
    RMCanvas.SceneFrame.AddRotation D3DRMCOMBINE_AFTER, Rnd * 2, Rnd * 2, Rnd * 2, 0.1
    RMCanvas.CameraFrame.SetColorRGB Rnd * 256, Rnd * 255, Rnd * 123
    Dim Superficie As DDSURFACEDESC2
    RMCanvas.DDraw.CreateSurfaceFromFile App.Path & "\cloud3.bmp", Superficie
    RMCanvas.DDraw.CreateSurface Superficie
End Sub

Private Sub Form_Unload(Cancel As Integer)
    M_Volando = False
End Sub

Private Sub mnu_Accidente_Click()
    mnu_Accidente.Checked = Not mnu_Accidente.Checked
End Sub

Private Sub mnu_Anima_Manual_Click()
If mnu_Anima_Manual.Checked Then Exit Sub
    mnu_Anima_Manual.Checked = True
    mnu_Auto_Anima.Checked = False
End Sub

Private Sub mnu_Auto_Anima_Click()
If mnu_Auto_Anima.Checked Then Exit Sub
    mnu_Anima_Manual.Checked = False
    mnu_Auto_Anima.Checked = True
End Sub

Private Sub RMCanvas_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, False)
End Sub

Private Sub RMCanvas_SceneMove(delta As Single)
If mnu_Auto_Anima.Checked Then _
     Movimiento_Camara_Avi�n delta
End Sub

Private Sub Timer_Render_Timer()
    RMCanvas.Update
    DoEvents
    mnu_FPS.Caption = RMCanvas.FPS
End Sub
