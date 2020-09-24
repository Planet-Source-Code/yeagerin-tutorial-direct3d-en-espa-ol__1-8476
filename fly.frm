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
      Caption         =   "&Animación"
      Begin VB.Menu mnu_Auto_Anima 
         Caption         =   "&Automática"
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
'Modelo del avión...
Dim M_Avión As Direct3DRMFrame3
'Cabina [ para controlar la posición el avión ]
Dim M_Cabina As Direct3DRMFrame3
'Su animación...
Dim M_AnimaAvión As Direct3DRMAnimation2
'Tiempo para la animación...
Dim M_Tiempo As Single
'¿Está volando?
Dim M_Volando As Boolean
Dim VelRot  'Velocidad Rotación
'Aclaraciones:
'======================================================================
'Un frame contiene la geometría con posiciones y orientación diferentes
'Un MeshBuilder es pura geometría y pueden cambiar la posición y orientación
' de sus objetos
'A su vez un frame puede ser vinculado a otro (una escena) y cuando se mueve
' todos los frames se mueven también junto con los mesh-builders añadidos a él.
'Es decir que cuando un frame se mueve todos sus mesh y frames tb lo hacen.
'--=[Es algo así como el sistema solar]=--------------------------------- :]
'  Donde el sol es el frame raíz (base),añadido a él se encuentran los
'frames childs o hijos que puede contener a la tierra, por ejemplo
' A su vez, añadido a la tierra (de geometría inferior) se encuentran añadidos
'otros frames hijos como la luna con su propia geometría, claro
'imagina que rotamos el Sol; esta rotación generaría en el frame de la tierra
'rotar alrededor de la órbita del frame Sol
' Pero podemos rotar la tierra y todo lo que ella contiene dentro del frame del Sol
'==================================================================================
'  RMCanvas.SceneFrame
'Este objeto es la base de todos los demás, incluido el raíz o base root
' es donde generamos nuestro mundo 3D
'---------
'  RMCanvas.CameraFrame
'   Este objeto es un hijo de la escena.
' Como su propio nombre indica determina la posición y orientación
'de la cámara.
' Su valor por defecto es -10 unidades detrás del eje Z y mirando
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
' Puedes usar el método setColorRGB para cambiar el color de la luz.
'---------
'  RMCanvas.AmbientLight
' Este determina como es la luz ambiental de la escena.
' * setColorRGB cambia el color y la intesidad del light
'   - no uses el blanco porque es "cegador", usa el gris claro :p
'---------
' RMCanvas.Viewport
' Esto describe como funciona la cámara:
'  * Usa el método setField para determinar como es de cerrada y amplia
' es el area que estás observando.
'  * Con setFront y setBack determinas desde cuánta distancia te gustaría
' ver y cómo de cerrado observas un objecto.
' Sirve además para obtener objetos que se están visualizando, incluso
'se le puede pasar el evento Mouse_Over del ratón... :]
'---------
'  RMCanvas.Device
'   Esto sirve para declarar el interfaz que controlará el rendering
' así como la calidad de dibujo (SetQuality).
' Lo normal es el Gouraund.
'---------
'  RMCanvas.SceneSpeed
'  En unidades por segundo, pueden ser ajustadas = que la rotación
' y la velocidad que afectan a un objecto.Default=30 unidadess/sg.
'---------
'=================================================
'En fin, estas son las propiedades más interesantes del control
' RMCanvas paraDirect3D, otras properties te permitirán por ejemplo
' dibujar sobre la escena como si fueran etiquetas, algo así como las
' animaciones del CounterStrike para Half-Life :]
'Mirad por ejemplo: DDBackSurface y sus ejemplos.
'=================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Not mnu_Anima_Manual.Checked Then Exit Sub
Dim Posición_Actual As D3DVECTOR
M_Avión.GetPosition M_Cabina, Posición_Actual
VelRot = VelRot + 0.0001
Select Case (KeyCode)
    Case 37: 'Izquierda... 37
        'M_Avión.AddRotation D3DRMCOMBINE_AFTER, Posición_Actual.X + 0.5, Posición_Actual.Y + 0.5, Posición_Actual.z + 0.5, -VelRot
        'M_Avión.AddRotation D3DRMCOMBINE_AFTER, 0#, 15#, 0#, VelRot
        
    Case 38: 'Arriba...
        M_Avión.AddRotation D3DRMCOMBINE_AFTER, 0#, 1#, 0#, 0.002
    Case 39: 'Derecha... 39
        M_Avión.AddRotation D3DRMCOMBINE_AFTER, 0#, -15#, 0#, VelRot
    Case 40: 'Abajo...
        M_Avión.SetPosition M_Cabina, Posición_Actual.X, Posición_Actual.Y + VelRot, Posición_Actual.z
        M_Avión.AddRotation D3DRMCOMBINE_AFTER, 0, 1#, 1#, VelRot
    Case 82: 'Rotar
        M_Avión.AddRotation D3DRMCOMBINE_AFTER, Posición_Actual.X + 0.5, Posición_Actual.Y + 0.5, Posición_Actual.z + 0.5, VelRot
    Case 76: 'Looping Star!
        M_Avión.AddRotation D3DRMCOMBINE_AFTER, Posición_Actual.X, Posición_Actual.Y, Posición_Actual.z + 10, VelRot
    Case Else:
        Exit Sub
End Select
M_AnimaAvión.SetFrame RMCanvas.CameraFrame
M_Avión.LookAt M_Cabina, Nothing, D3DRMCONSTRAIN_Z
'RMCanvas.CameraFrame.SetPosition Nothing, Posición_Actual.x - 10, Posición_Actual.y - 10, Posición_Actual.z - 10
RMCanvas.CameraFrame.LookAt M_Avión, Nothing, D3DRMCONSTRAIN_Z
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    VelRot = 0.002
End Sub

Private Sub Form_Load()
Static Rula As Boolean
'Creamos una variable estática para llamadas entre módulos
' es decir que no pierda su valor al hacer referencia a él
' desde otra función y sepamos si rulaba o no el 3D :]
    If Rula = True Then End
   Rula = True
'Mostrar el formulario antes de cargar la escena
    Me.Show
'Hacer los demás eventos antes de la escena...
    DoEvents
    VelRot = 0.0002
    Inicio
End Sub
Private Sub Form_Resize()
'Tamaño del Canvas al mismo tamaño que la ventana...
    RMCanvas.width = Me.ScaleWidth
    RMCanvas.height = Me.ScaleHeight
End Sub
Sub Inicio()
Dim Rula As Boolean
Dim sFile As String
'  Vamos a ejecutarlo en modo ventana...
'  encontramos el hardware 3D del 'Display' Primario...
'  si no lo encuentra usará el Software emulation RGB rasterizer
Rula = RMCanvas.StartWindowed
    If Rula = False Then
        MsgBox "¡No puedo iniciar Direct3D con Hardware!" & vbCrLf & _
        vbCrLf & " · Renderizando con software...", vbCritical, "Error de Hardware"
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
    Crear_Trayectoria_Animación
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
'- Y cargamos su geometría de un ficherofile en el constructor de estructuras
' añadiéndolo en el frame.
    mb_deTierra.LoadFromFile App.Path & "\land4.x", 0, 0, Nothing, Nothing
    m_tierra.AddVisual mb_deTierra
'Hacemos que la tierra sea mayor escalándola y tomando su extensión
' la colocamos en una caja, así sabremos cuánto mide exactamente, es como
' tomar las medidas de una fase de un juego para que el personaje
' no se salga de ella. :)
    mb_deTierra.ScaleMesh 10, 8, 10
    mb_deTierra.GetBox Caja
'Metemos sus medidas en la caja, un poco de imaginación :p
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
'Dibujamos las caras según su largo:
    Dim Vertice As D3DVECTOR, Vector_Normal As D3DVECTOR, Y As Single
'Recorremos todos los vertices del MeshBuilder de Tierra
    For i = 0 To mb_deTierra.GetFaceCount() - 1
        Y = Caja.Min.Y
    'Y a su vez por todas las caras...
        For j = 0 To mb_deTierra.GetFace(i).GetVertexCount() - 1
            mb_deTierra.GetFace(i).GetVertex j, Vertice, Vector_Normal
    'Colocamos todos los vértices dentro de la caja
            If Vertice.Y > Y Then Y = Vertice.Y
        Next
        If (Y - Caja.Min.Y) / RangO < 0.05 Then
    'Si la cara por la que vamos en el bucle se ve por encima de la tierra, es decir
    ' la cámara lo ve se dibuja de color potito sino pues de blanco,,,
            Call mb_deTierra.GetFace(i).SetColorRGB((Y - Caja.Min.Y) / RangO, 0.6, 1 - (Y - Caja.Min.Y) / RangO)
        Else
            Call mb_deTierra.GetFace(i).SetColorRGB(0.2 + (Y - Caja.Min.Y) / RangO, 1 - (Y - Caja.Min.Y) / RangO, 0.5)
        End If
    Next
End Sub

Sub Crear_Modelo()
    Dim mb_Avión As Direct3DRMMeshBuilder3
'Creamos un MeshBuilder Avión para dibujar el modelo...
    Set mb_Avión = RMCanvas.D3DRM.CreateMeshBuilder()
'Cargandolo de un fichero.X...
    mb_Avión.LoadFromFile App.Path & "\dropship.x", 0, 0, Nothing, Nothing
'Le ajustamos el tamaño...
    mb_Avión.ScaleMesh 0.015, 0.008, 0.015
    ' y el color...
    mb_Avión.SetColorRGB 0.8, 0.8, 0.8
'Creamos un Frame para representar el modelo...
Set M_Avión = RMCanvas.D3DRM.CreateFrame(RMCanvas.SceneFrame)
  ' añadiéndolo a la escena
    M_Avión.AddVisual mb_Avión
    Dim Textura As Direct3DRMTexture3
    Set Textura = RMCanvas.CreateUpdateableTexture(64, 64, App.Path & "\banana.bmp")
    Textura.GenerateMIPMap
    Textura.SetName "Banana"
    M_Avión.SetTexture Textura
    M_Avión.GetParent.SetTexture Textura
'Lo mismo con su chase (cabina)...
Set M_Cabina = RMCanvas.D3DRM.CreateFrame(RMCanvas.SceneFrame)
'Crea un array de 1000 vértices de DirectX ¿ :P ?
'Dim verts(1000) As D3DRMVERTEX
End Sub

Sub Crear_Trayectoria_Animación()
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
' crear la animación:
Set M_AnimaAvión = RMCanvas.D3DRM.CreateAnimation()
'Las opciones de la animación están establecidas aquí abajo para que
' se repita continuamente...
    M_AnimaAvión.SetOptions D3DRMANIMATION_CLOSED Or D3DRMANIMATION_SPLINEPOSITION Or D3DRMANIMATION_POSITION
Dim Posición As D3DRMANIMATIONKEY
'Este bucle va de 10 en 10 cambiando la posición del modelo:
    For i = 0 To 10
'Toma los datos escalares del array de posiciones preparado
' especialmente para este "mapa".
        X = Datos_Trayectoria(i * 3)
        Y = Datos_Trayectoria(i * 3 + 1)
        z = Datos_Trayectoria(i * 3 + 2)
'Esto de aquí abajo se ahorra con: m_AnimaAvión.AddPositionKey i, x, y, z
        Posición.dvX = X
        Posición.dvY = Y
        Posición.dvZ = z
        Posición.lKeyType = 3
        Posición.dvTime = i
'Añadimos el juego de posiciones a la animación para que sepa donde
' tiene que colocar el objeto.
        M_AnimaAvión.AddKey Posición
    Next
End Sub

Sub Movimiento_Camara_Avión(delta As Single)
    Dim Dirección As D3DVECTOR
    Dim Dirección_Antigüa As D3DVECTOR
    Dim Dirección_Cámara As D3DVECTOR
    Dim Dir_Ant_Cam As D3DVECTOR
'Velocidad de la escena...
    RMCanvas.SceneSpeed = 1
'El tiempo de la animación va cambiando con el movimiento de
' la cámara el delta es el valor de la propia escena...
    M_Tiempo = M_Tiempo + delta
'Colocamos la cámara según la animación:
    M_AnimaAvión.SetFrame RMCanvas.CameraFrame
    M_AnimaAvión.SetTime M_Tiempo + 0
'...al igual que el modelo del avión...
    M_AnimaAvión.SetFrame M_Avión
    M_AnimaAvión.SetTime M_Tiempo + 0.5
'...y el chase (cabina)...
    M_AnimaAvión.SetFrame M_Cabina
    M_AnimaAvión.SetTime M_Tiempo + 1
'orientamos la cámara hacia el avión...
    RMCanvas.CameraFrame.LookAt M_Avión, Nothing, D3DRMCONSTRAIN_Z
'y el avión a su vez lo orientamos hacia la cabina...
    M_Avión.LookAt M_Cabina, Nothing, D3DRMCONSTRAIN_Y
'tomamos la orientación de la cámara...
    RMCanvas.CameraFrame.GetOrientation Nothing, Dirección_Antigüa, Dir_Ant_Cam
'y la del avión...
    M_Avión.GetOrientation Nothing, Dirección, Dirección_Antigüa
'almacenándolas en sus correspondientes variables...
    Dirección_Antigüa.X = Dirección.X - Dirección_Antigüa.X
    Dirección_Antigüa.Y = Dirección.Y - Dirección_Antigüa.Y + 1#
    Dirección_Antigüa.z = Dirección.z - Dirección_Antigüa.z
'Ahora podemos colocar al avión (:] eehhee) en su sitio:
    M_Avión.SetOrientation Nothing, Dirección.X, Dirección.Y, Dirección.z, Dirección_Antigüa.X, Dirección_Antigüa.Y, Dirección_Antigüa.z
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
     Movimiento_Camara_Avión delta
End Sub

Private Sub Timer_Render_Timer()
    RMCanvas.Update
    DoEvents
    mnu_FPS.Caption = RMCanvas.FPS
End Sub
