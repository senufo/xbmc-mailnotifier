import xbmc, xbmcgui
import xbmcaddon
import poplib
import time
#Traitement des fichiers xml
import xml.dom.minidom

__author__     = "Senufo"
__scriptid__   = "service.notifier"
__scriptname__ = "Notifier"

Addon      = xbmcaddon.Addon(id=__scriptid__)

__cwd__        = Addon.getAddonInfo('path')
__version__    = Addon.getAddonInfo('version')
__language__   = Addon.getLocalizedString

__profile__    = xbmc.translatePath( Addon.getAddonInfo('profile') )
__resource__   = xbmc.translatePath( os.path.join( __cwd__, 'resources', 'lib' ) )

sys.path.append (__resource__)


print "dirHome : %s " % __cwd__
#ID de la fenetre HOME
WINDOW_HOME = 10000 
#print "==============NOTIFIER==============="
NbMsg = [0,0,0]
#POPHOST = "192.168.10.254"
#Verifie que xbmc tourne
while (not xbmc.abortRequested):
	  #On vide le message
          msg = ''
	  NxMsgTot = 0
	  MsgTot = False
	  #On recupere les parametres des trois serveurs
	  for i in [1,2,3]:
		id = 'user' + str(i)
	   	USER = Addon.getSetting( id )
		id = 'name' + str(i)
    		NOM =  Addon.getSetting( id )
		id = 'server' + str(i)
		SERVER = Addon.getSetting( id )
		id = 'pass' + str(i)
		PASSWORD =  Addon.getSetting( id )
		id = 'port' + str(i)
		PORT =  Addon.getSetting( id )
		id = 'ssl' + str(i)
		SSL = Addon.getSetting( id )
#		print "SERVER = %s, PORT = %s, USER = %s, password = %s, SSL = %s" % (SERVER,PORT,USER, PASSWORD, SSL)
                #Teste si USER existe 
		if (USER != ''):
		   try :
			SSL.lower()
			if ('no' in SSL): 
				mail = poplib.POP3(str(SERVER),int(PORT))
			else: 
				mail = poplib.POP3_SSL(str(SERVER),int(PORT))
		        mail.user(str(USER))
		        mail.pass_(str(PASSWORD))
			
	        	numEmails = mail.stat()[0]
			if numEmails > 0: MsgTot = True
			#On regarde si un nouveau mail est arrive
			if NbMsg[i] == 0: NbMsg[i] = numEmails
			NxMsg = numEmails - NbMsg[i]
			if NxMsg > 0 : 
				NbMsg[i] = NbMsg[i] + NxMsg
				NxMsgTot = NxMsgTot + NxMsg
			else:
				NbMsg[i] = numEmails
		        #print "numEmails = %d " % numEmails
			locstr = Addon.getLocalizedString(id=610)
			if numEmails != 0:
		                msg = msg + "%s : %d %s" % (NOM,numEmails, locstr) + "\n"
			numEmails = 0
	           except:
			locstr = Addon.getLocalizedString(id=613)
                        xbmc.executebuiltin("XBMC.Notification(%s :, %s,30)" % (locstr,SERVER))
		        print "Erreur de connection : %s" % SERVER
		        msg = msg + "%s : %s\n" % (NOM,locstr)
	  homeWin = xbmcgui.Window(WINDOW_HOME)
          x = int(Addon.getSetting( 'x' ))
          y = int(Addon.getSetting( 'y' ))
          height = int(Addon.getSetting( 'height' ))
          width = int(Addon.getSetting( 'width' ))
	  print "POS =%d, %d, %d, %d " % (x,y,height,width)
	  #On n'affiche que si il y a des messages
	  if MsgTot:
	          #Affiche l'icone de mail
		  ImgBox = xbmcgui.ControlImage(x-75, y, 75, 75, __cwd__ + '/home-myemail.png') #aspectRatio=1, , colorDiffuse='0xC0FF0000'
		  font_msg = Addon.getSetting( 'font' )
		  couleur_msg = Addon.getSetting( 'color' )
		  MsgBox = xbmcgui.ControlTextBox(x, y, width, height, font=font_msg, textColor=couleur_msg)
		  print "Message = %s " % msg
		  homeWin.addControl(ImgBox)
		  homeWin.addControl(MsgBox)

		  message =  "%s " % msg
	 	  MsgBox.setText(message)
	  if NxMsgTot > 0:
		locstr = Addon.getLocalizedString(id=611)
		xbmc.executebuiltin("XBMC.Notification( ,%d %s,60)" % (NxMsgTot,locstr))

	  #Attente avant de relever les mails
	  intervalle = Addon.getSetting( 'time' )
	  time.sleep(int(intervalle))
	  #Efface la textbox si elle existe
	  try:
		homeWin.removeControl(MsgBox)
	  except:
                print "Le controle n'existe pas"
        
