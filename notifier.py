import xbmc, xbmcgui
import xbmcaddon
import poplib, imaplib
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


#ID de la fenetre HOME
WINDOW_HOME = 10000 
NbMsg = [0,0,0,0]
numEmails = 0 
#ID du screen HOME
homeWin = xbmcgui.Window(WINDOW_HOME)
currentWindowId = xbmcgui.getCurrentWindowId()
#On stocke le temps courant
temps = time.clock()
#On prend l'intervalle entre deux releve de boite
intervalle = int(Addon.getSetting( 'time' ))
#On stocke le temps precedent, pour la 1er fois
#on enleve intervalle pour un affichage des le debut
old_temps = temps - intervalle
#Total des messages, nx messages, contenu de l'affichage
NxMsgTot = 0
msg = ''
message = ''
MsgTot = True
#No du serveur
NoServ = 1

def Checkmail(USER,NOM,SERVER,PASSWORD,PORT,SSL,TYPE,MsgTot,NoServ):
		print "CHECKMAIL"
		msg = ''
                NxMsgTot = 0
                numEmails = 0 
		i = NoServ
		print "Ligne 55"
                #Teste si USER existe 
		if (USER != ''):
		   try:
			#Partie POP3
			if  '0' in TYPE:  #'POP' 
			  if SSL:
			       mail = poplib.POP3_SSL(str(SERVER),int(PORT))
			  else:  #'POP3'
			       mail = poplib.POP3(str(SERVER),int(PORT))
                          mail.user(str(USER))
		          mail.pass_(str(PASSWORD))
	        	  numEmails = mail.stat()[0]
                        #Partie IMAP
                        if '1' in TYPE:
			  if SSL:
			       imap = imaplib.IMAP4_SSL(SERVER, int(PORT))
	                  else:
		               imap = imaplib.IMAP4(SERVER, int(PORT))
	  		  imap.login(USER, PASSWORD)
                       	  id = 'folder' + str(i)
		          FOLDER = Addon.getSetting( id )
		          imap.select(FOLDER)
	                #  numEmails = 1 
	                  numEmails = len(imap.search(None, 'UnSeen')[1][0].split()) 
		          #print "IMAP numEmails = %d " % numEmails
                   except:
		   	locstr = Addon.getLocalizedString(id=613) #Erreur de connection
			if Addon.getSetting( 'erreur' ) == "true":
                        	xbmc.executebuiltin("XBMC.Notification(%s :, %s,30)" % (locstr,SERVER))
		        #print "Erreur de connection : %s, erreur = %s" % (SERVER, mail)
		        msg = msg + "%s : %s\n" % (NOM,locstr)

		   if numEmails > 0: MsgTot = True #Il y a des messages
		   #On regarde si un nouveau mail est arrive
	           if NbMsg[i] == 0: NbMsg[i] = numEmails
		   NxMsg = numEmails - NbMsg[i]
		   if NxMsg > 0 : 
			NbMsg[i] = NbMsg[i] + NxMsg
			NxMsgTot = NxMsgTot + NxMsg
	           else:
			NbMsg[i] = numEmails
		   #print "numEmails = %d " % numEmails
		   locstr = Addon.getLocalizedString(id=610) #messages(s)
		   if numEmails != 0:
		          #msg = msg + "%s : %d %s" % (NOM,numEmails, locstr) + "\n"
		          msg = msg + "%s : %d " % (NOM,numEmails) + "\n"
		          numEmails = 0
                   if NxMsgTot > 0:
	       	      locstr = Addon.getLocalizedString(id=611) #Nouveau(x) message(s)
                      #Affiche un popup a l'arrivee d'un nx msg
		      if Addon.getSetting( 'popup' ) == "true":
		            xbmc.executebuiltin("XBMC.Notification( ,%d %s,60)" % (NxMsgTot,locstr))
	           return MsgTot,msg


#Verifie que xbmc tourne
while (not xbmc.abortRequested):
	  #Attente avant de relever les mails
	  #On ne releve les mails qui si on a attendu
	  #un temps >  a l'intervalle
          temps =  time.clock() 
	  if (temps > (old_temps + intervalle)):
            #print "temps_sup = %f, old_temps = %f" % (temps, old_temps)
	    old_temps = temps
	    #print "time %f " % temps
            #On vide le message
            msg = ''
	    message = ''
	    #Total des nx messages
	    NxMsgTot = 0
	    #Pas de nx message
	    MsgTot = False
	    #Quel type d'affichage : tout les serveurs a la fois ou l'un apres l'autre
            ALT = Addon.getSetting( 'alt' ) == "true"
	    #Si l'un apres l'autre
	    if ALT:
	     USER = Addon.getSetting( 'user' + str(NoServ) )
    	     NOM =  Addon.getSetting( 'name' + str(NoServ))
	     SERVER = Addon.getSetting( 'server' + str(NoServ))
	     PASSWORD =  Addon.getSetting( 'pass' + str(NoServ))
	     PORT =  Addon.getSetting( 'port' + str(NoServ))
	     SSL = Addon.getSetting( 'ssl' + str(NoServ)) == "true"
	     TYPE = Addon.getSetting( 'type' + str(NoServ))
	     print "=>SERVER = %s, PORT = %s, USER = %s, password = %s, SSL = %s, TYPE = %s" % (SERVER,PORT,USER, PASSWORD, SSL,TYPE)
             if (USER != ''):
               (MsgTot, msg) = Checkmail(USER,NOM,SERVER,PASSWORD,PORT,SSL,TYPE, MsgTot,NoServ)
	     message = message + msg
     	     NoServ += 1  #On passe au serveur suivant
	     if (NoServ > 3): NoServ =1 #Si dernier serveur on revient au premier
	    else:  #Affichage de tout les serveurs a la fois
	     #On recupere les parametres des trois serveurs
	     for i in [1,2,3]:
	   	USER = Addon.getSetting( 'user' + str(i) )
    		NOM =  Addon.getSetting( 'name' + str(i) )
		SERVER = Addon.getSetting( 'server' + str(i) )
		PASSWORD =  Addon.getSetting( 'pass' + str(i) )
		PORT =  Addon.getSetting( 'port' + str(i) )
		SSL = Addon.getSetting( 'ssl' + str(i) ) == "true"
		TYPE = Addon.getSetting( 'type' + str(i) )
		#print "SERVER = %s, PORT = %s, USER = %s, password = %s, SSL = %s, TYPE = %s" % (SERVER,PORT,USER, PASSWORD, SSL,TYPE)
                if (USER != ''):
		   (MsgTot, msg) = Checkmail(USER,NOM,SERVER,PASSWORD,PORT,SSL,TYPE, MsgTot,i)
		message = message + msg


            #On n'affiche que si il y a des messages
	    #Verifie si chgt de Home SCREEN
            if MsgTot:
                if (xbmcgui.getCurrentWindowId() == WINDOW_HOME):
		 #666 = id du control Label dans le Home.xml
                 MsgBox = homeWin.getControl(666)
                 MsgBox.setLabel(message)
		 #On efface le contenu
		 message = ''
