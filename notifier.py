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
#ID du screen HOME
homeWin = xbmcgui.Window(WINDOW_HOME)
currentWindowId = xbmcgui.getCurrentWindowId()
windowChange = False
print "homeWin ", homeWin
print "curwin ", currentWindowId
#On stocke le temps courant
temps = time.clock()
#On prend l'intervalle entre deux releve de boite
intervalle = int(Addon.getSetting( 'time' ))
#On stocke le temps precedent, pour la 1er fois
#on enleve intervalle pour un affichage des le debut
old_temps = temps - intervalle
#Verifie que xbmc tourne
while (not xbmc.abortRequested):
	  #Attente avant de relever les mails
	  #On ne releve les mails qui si on a attendu
	  #un temps >  a l'intervalle
          #print "temps_avt = %f, old_temps = %f" % (temps, old_temps)
          temps =  time.clock() 
	  if (temps > (old_temps + intervalle)):
            #print "temps_sup = %f, old_temps = %f" % (temps, old_temps)
	    old_temps = temps
	    #print "time %f " % temps
            #On vide le message
            msg = ''
	    #Total des nx messages
	    NxMsgTot = 0
	    #Pas de nx message
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
		   try:
			SSL.lower()
			if ('no' in SSL): 
				mail = poplib.POP3(str(SERVER),int(PORT))
			else: 
				mail = poplib.POP3_SSL(str(SERVER),int(PORT))
		        mail.user(str(USER))
		        mail.pass_(str(PASSWORD))
			
	        	numEmails = mail.stat()[0]
			if numEmails > 0: MsgTot = True #Il y a des messages
			#On regarde si un nouveau mail est arrive
			if NbMsg[i] == 0: NbMsg[i] = numEmails
			NxMsg = numEmails - NbMsg[i]
			if NxMsg > 0 : 
				NbMsg[i] = NbMsg[i] + NxMsg
				NxMsgTot = NxMsgTot + NxMsg
			else:
				NbMsg[i] = numEmails
		        print "numEmails = %d " % numEmails
			locstr = Addon.getLocalizedString(id=610) #messages(s)
			if numEmails != 0:
		                msg = msg + "%s : %d %s" % (NOM,numEmails, locstr) + "\n"
			numEmails = 0
	           except:
			locstr = Addon.getLocalizedString(id=613) #Erreur de connection
                        xbmc.executebuiltin("XBMC.Notification(%s :, %s,30)" % (locstr,SERVER))
		        print "Erreur de connection : %s" % SERVER
		        msg = msg + "%s : %s\n" % (NOM,locstr)
	    if NxMsgTot > 0:
		locstr = Addon.getLocalizedString(id=611) #Nouveau(x) message(s)
		xbmc.executebuiltin("XBMC.Notification( ,%d %s,60)" % (NxMsgTot,locstr))

 #On n'affiche que si il y a des messages
	    #Verifie si chgt de Home SCREEN
            if MsgTot:
                if (xbmcgui.getCurrentWindowId() == WINDOW_HOME):
		 #666 = id du control Label dans le Home.xml
                 MsgBox = homeWin.getControl(666)
                 MsgBox.setLabel(msg)

	         
