VERSION 5.00
Object = "{603D6079-7088-48DB-9688-A354A8BA98AA}#3.0#0"; "MathEqu.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Form"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MathEQ.ctlMathDraw ctlFormular 
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7646
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      FontName        =   "Schoolnet Sans Serif"
      FontSize        =   9.75
      LineSpace       =   0
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Render"
      Default         =   -1  'True
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txtContent 
      BeginProperty Font 
         Name            =   "Schoolnet Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4440
      Width           =   9615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Double click on formula to edit"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   7530
      Width           =   2145
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  ctlFormular.Text = txtContent.Text
  ctlFormular.RenderOnScreen
End Sub

Private Sub ctlFormular_ContentChanged()
  txtContent.Text = ctlFormular.Text
End Sub

Private Sub Form_Load()
  Dim st As String
  Dim i As Long
  st = "C¸c phÐp to¸n c¬ b¶n: a+b-(c/(d*e))^2+((a+b)^2)-+3/(x+2)" & vbNewLine
  st = st & "PhÐp lòy thõa: ax^2+bx+c=0" & vbNewLine
  st = st & "ChØ sè d­íi: a_i+b_j=c_k" & vbNewLine
  st = st & "C¨n bËc hai: x_12=(-b+-sqrt(b^2-4ac))/2a" & vbNewLine
  st = st & "PhÐp nèi: {Comb(2,1/3)+Comb(3,1/4)+Comb(4,1/5)}" & vbNewLine
  st = st & "DÊu tæng: {Sum(x_i/(x_i+1),i=1,n)}" & vbNewLine
  st = st & "DÊu tÝch: {Prod(x_i+1/x_i, i=1, n)}" & vbNewLine
  st = st & "DÊu hîp: {Uni(x_i,i=1,n)}" & vbNewLine
  st = st & "DÊu tÝch ph©n: {int(x/(2+x),dx)} vµ {int(x/(2+x),dx,1,10/(x+1))}" & vbNewLine
  st = st & "C¸c ký hiÖu Hy L¹p: {(2&alpha;+3&beta;)/&mu;=10&epsilon;} hoÆc {(2*&alpha;+3*&beta;)/&mu;=10^&epsilon;}" & vbNewLine
  st = st & "alpha, beta, chi, delta, epsilon, phi, phiv, gamma, eta, kappa, lamda, mu, nu, pi, piv, theta, rho, sigma, finalsigma, tau, upsilon, omega, xi, pxi, zeta, omicron: {comb(&alpha;, &beta;, &chi;, &delta;, &epsilon;, &phi;, &phiv;, &gamma;, &eta;, &kappa;, &lamda;, &mu;, &nu;, &pi;, &piv;, &theta;, &rho;, &sigma;, &finalsigma;, &tau;, &upsilon;, &omega;, &xi;, &pxi;, &zeta;, &omicron;,&infinity;)}" & vbNewLine
  st = st & "{comb(&Alpha;, &Beta;, &Chi;, &Delta;, &Epsilon;, &Phi;, &Gamma;, &Eta;, &Iota;, &Kappa;, &Lamda;, &Mu;, &Nu;, &Pi;, &Theta;, &Rho;, &Sigma;, &Tau;, &Upsilon;, &Omega;, &Xi;, &Psi;, &Zeta;, &Omicron;)}" & vbNewLine
  st = st & "BiÓu thøc l­îng gi¸c: {tg(&alpha;)=sin(&alpha;)/cos(&alpha;)}" & vbNewLine
  st = st & "CÊu t¹o sè: {over(abc)=100a+10b+c}" & vbNewLine
  st = st & "Ma trËn: {matrix(3,3,a_11,a_12,a_13,a_21,a_22,a_23,a_31,a_32,a_33)} vµ §Þnh thøc: {det(3,3,a_11,a_12,a_13,a_21,a_22,a_23,a_31,a_32,a_33)}" & vbNewLine
  st = st & "HÖ ph­¬ng tr×nh: {equ(3, comb(a_1,x)+comb(b_1,y)+comb(c_1,z)=0,comb(a_2,x)+comb(b_2,y)+comb(c_2,z)=0,comb(a_3,x)+comb(c_3,z)=0)}" & vbNewLine
  st = st & "Vector: {vector(AB)+vector(BC)=vector(AC)}" & vbNewLine
  st = st & "Gi¸ trÞ tuyÖt ®èi: {abs(a)} hoÆc {abs(A)=det(3,3,a_11,a_12,a_13,a_21,a_22,a_23,a_31,a_32,a_33)}" & vbNewLine
  st = st & "Ký hiÖu gãc: {ang(ABC)+ang(ACB)+ang(BAC)=180^o}" & vbNewLine
  st = st & "Ký hiÖu cung: {arc(AB)+arc(BC)=arc(ABC)+arc(CD)=arc(ABCD)}" & vbNewLine
  st = st & "C¸c phÐp to¸n quan hÖ: {a+b<c/e>=a^2+1<=(x+y)/2>1/3<>a/b}" & vbNewLine
  st = st & "C¸c dÊu mòi tªn: {comb(&larrow;,&rarrow;,&uarrow;,&darrow;,&lrarrow;,&larrowd;,&rarrowd;,&uarrowd;,&darrowd;,&lrarrowd;,&urarrow;,&drarrow;)}" & vbNewLine
  st = st & "{comb(a^2+b^2=c^2,&rarrowd;)} a,b,c lµ ba c¹nh cña tam gi¸c vu«ng" & vbNewLine
  st = st & "{comb((a+b)^2,&lrarrowd;,a^2+comb(2,a,b)+b^2)}" & vbNewLine
  st = st & "Ký hiÖu h×nh häc: {comb(AB,&perp;,AC)}, {comb(AB,&parallel;CD)}" & vbNewLine
  st = st & "Ký hiÖu tËp hîp: víi {comb(&any;,x,&in;,A)}, ta cã {comb(&exist;,x_i,:,x_i+1=x_i+x_(i-1))}" & vbNewLine
  st = st & "Ký hiÖu giíi h¹n: {lim(x/(1+x^2),comb(x,&rarrow;,&infinity;))}" & vbNewLine
  st = st & "HÖ ph­¬ng tr×nh víi sè thø tù: {nmatrix(1,2,equ(3, comb(a_1,x)+comb(b_1,y)+comb(c_1,z)=0,comb(a_2,x)+comb(b_2,y)+comb(c_2,z)=0,comb(a_3,x)+comb(c_3,z)=0),(1))}" & vbNewLine
  st = st & "Mét sè ký hiÖu kh¸c: {Comb(&subset;,&supset;,&vdots;,&empty;,&lor;,&land;,&sim;,&simeq;,&ll;,&gg;,&approx;,&propto;,&equiv;)}"
  txtContent.Text = st
End Sub

