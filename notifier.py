import xbmc, xbmcgui
import xbmcaddon
import poplib, imaplib
import time

__author__     = "Senufo"
__scriptid__   = "service.notifier"
__scriptname__ = "Notifier"

Addon          = xbmcaddon.Addon(__scriptid__)

__cwd__        = Addon.getAddonInfo('path')
__version__    = Addon.getAddonInfo('version')
__language__   = Addon.getLocalizedString

__profile__    = xbmc.translatePath(Addon.getAddonInfo('profile'))
__resource__   = xbmc.translatePath(os.path.join(__cwd__, 'resources',
                                                 'lib'))

DEBUG_LOG = Addon.getSetting( 'debug' )
if 'true' in DEBUG_LOG : DEBUG_LOG = True
else: DEBUG_LOG = False

#sys.path.append (__resource__)
#Function Debug
def debug(msg):
    """
    print message if DEBUG_LOG == True
    """
    if DEBUG_LOG == True: print " [%s] : %s " % (__scriptname__, msg)

#ID de la fenetre HOME
WINDOW_HOME = 10000

#Variables generales
msg = ''
NbMsg = [0, 0, 0, 0]
numEmails = 0
#No serveur
NoServ = 1
#Position du texte
x = int(Addon.getSetting('x'))
y = int(Addon.getSetting('y'))
width = int(Addon.getSetting('width'))
height = int(Addon.getSetting('height'))
font = Addon.getSetting('font')
color = Addon.getSetting('color')
#Affichage alternatif ??
ALT = Addon.getSetting('alt')
#Display in the skin
SKIN = Addon.getSetting('skin')
#ID du controle dans la fenetre Home.xml
MsgBox = None
MsgBoxId = None

start_time = 0
#Flag for add or not control in HOME
re_added_control = False
#Verifie que xbmc tourne
while (not xbmc.abortRequested):
    #Attente avant de relever les mails
    intervalle = int(float(Addon.getSetting('time')) * 60.0)
    if start_time and (time.time() - start_time) < intervalle:
        time.sleep(.5)
        SHOW_UPDATE     = Addon.getSetting('show_update') == "true"
        #if control exist
        if MsgBox:
            # optional show le temps qu'il reste avant la prochaine MAJ
            try:
                #If SHOW_UPDATE true
                if SHOW_UPDATE:
                    up = int(intervalle) - (time.time() - start_time)
                    locstr = Addon.getLocalizedString(615)  #Update in %i second
                    debug( "MSG up = %s " % msg)
                    label = "%s[CR], %s : %s" %  (msg,locstr,up)
                    #debug_string = "Msg = %s, Update = %s" %  (msg, up)
                    #debug( "label = %s " % debug_string)
                else: #Il faut rafraichir l'affichage
                    debug( "MSG = %s " % msg)
                    label = '%s' % msg
                if (SKIN == "false"):
                    MsgBox.setLabel( msg )
                    debug( "setlabel : %s" % msg)
                else:
                    MsgBox.setLabel( '' )
                    debug( 'Clean label')
            except Exception, e:
                print str(e)
        #Fin du if MsgBox
        HomeNotVisible = xbmc.getCondVisibility( "!Window.IsVisible(10000)" )
        if HomeNotVisible:
            #oop! on est plus sur le home
            re_added_control = True
        #elif re_added_control and not HomeNotVisible:
        else:
            #Try to get getcontrol if not exist make a new one
            try:
                MsgBox = homeWin.getControl( MsgBoxId )
            except:
                MsgBox = xbmcgui.ControlLabel( x, y, width, height, '', font, color )
            # add control label and set default label
            try:
                homeWin.addControl( MsgBox )
            except:
                pass
            # get control id
            MsgBoxId = MsgBox.getId()
            #Not used now ?
            re_added_control = False
            # reload addon setting possible change
            Addon = xbmcaddon.Addon( __scriptid__ )
        #Fin du HomeNotVisible
        # continue le while sans faire le reste
        continue
    #If firstime get ID of WINDOW_HOME
    homeWin = xbmcgui.Window(WINDOW_HOME)
    #Verif if control exist
    if MsgBoxId:
        try:
            MsgBox = homeWin.getControl( MsgBoxId )
        except:
            MsgBoxId = None
    #If no exist make a newone
    if MsgBoxId is None:
        MsgBox = xbmcgui.ControlLabel( x, y, width, height, '', font, color )
        #retire le control s'il exist # pas vraiment besoin le test a ete fait avec homeWin.getControl( MsgBoxId )
        try: homeWin.removeControl( MsgBox )
        except:
            debug("Le controle n\'existe pas")
            pass
        # add control label and set default label
        homeWin.addControl( MsgBox )
        # get control id
        MsgBoxId = MsgBox.getId()
    #Display update msg
    locstr = Addon.getLocalizedString(616) #Mise a jour
    MsgBox.setLabel( locstr % ' ' )

    #On vide le message
    msg = ''

    #On recupere les parametres des trois serveurs
    for i in range( 1, 4 ): #[1,2,3]:
        ENABLE =  Addon.getSetting( 'enableserver%i' % i )
        homeWin.setProperty( ("notifier.enable%i" % i) , ("%s" % ENABLE ))
        if ENABLE == "false":
            #homeWin.setProperty( ("notifier.enable%i" % i) , ("false"))
            #debug( :Enableserver = %s, i = %d  " % (Addon.getSetting(
            #    'enableserver%i' % i), i)
            #Si le serveur n'est pas defini on passe au suivant
            continue
        USER     = Addon.getSetting( 'user%i'   % i )
        NOM      = Addon.getSetting( 'name%i'   % i )
        SERVER   = Addon.getSetting( 'server%i' % i )
        PASSWORD = Addon.getSetting( 'pass%i'   % i )
        PORT     = Addon.getSetting( 'port%i'   % i )
        SSL      = Addon.getSetting( 'ssl%i'    % i )
        TYPE     = Addon.getSetting( 'type%i'   % i )
        FOLDER   = Addon.getSetting( 'folder%i' % i )

        #debug( :SERVER = %s, PORT = %s, USER = %s, password = %s, SSL = %s" % (SERVER, PORT, USER, PASSWORD, SSL)
#Total des nx messages
        NxMsgTot = 0
#Pas de nx message
        MsgTot = False

# Teste si USER existe
        if (USER != ''):
            try:
                locstr = Addon.getLocalizedString(616) #Releve du courrier
                MsgBox.setLabel( locstr % NOM )
#Partie POP3
                if  '0' in TYPE:  #'POP'
                    if SSL.lower() == 'false':
                        mail = poplib.POP3(str(SERVER), int(PORT))
                    else:
                        mail = poplib.POP3_SSL(str(SERVER), int(PORT))
                    mail.user(str(USER))
                    mail.pass_(str(PASSWORD))
                    numEmails = mail.stat()[0]
                    debug( "POP numEmails = %d " % numEmails)
#Partie IMAP
                if '1' in TYPE:
                    if SSL.lower() == 'true':
                        imap = imaplib.IMAP4_SSL(SERVER, int(PORT))
                    else:
                        imap = imaplib.IMAP4(SERVER, int(PORT))
                    imap.login(USER, PASSWORD)
                    FOLDER = Addon.getSetting( 'folder%i' % i )
                    imap.select(FOLDER)
                    numEmails = len(imap.search(None, 'UnSeen')[1][0].split())
                    debug( "IMAP numEmails = %d " % numEmails)

                #debug( :numEmails = %d " % numEmails
                locstr = Addon.getLocalizedString(610) #message(s)
                #msg = msg + "%s : %d %s" % (NOM,numEmails, locstr) + "\n"
                #numEmails = 0
            except:
                locstr = Addon.getLocalizedString(613) #erreur de connection
                if Addon.getSetting( 'erreur' ) == "true":
                    xbmc.executebuiltin("XBMC.Notification(%s : ,%s,30)" % (locstr, SERVER))
                debug( "Erreur de connection : %s" % SERVER)
#Msg affiche sur le HOME
            msg = msg + "%s : %s\n" % (NOM, locstr)

            if numEmails > 0:
                MsgTot = True #Il y a des messages
            #On regarde si un nouveau mail est arrive
            #dans NbMsg on stocke le nb de messages qui sont sur le serveur
            if NbMsg[i] == 0:
                NbMsg[i] = numEmails
            NxMsg = numEmails - NbMsg[i]
            if NxMsg > 0:
                NbMsg[i] = NbMsg[i] + NxMsg
                NxMsgTot = NxMsgTot + NxMsg
            else:
                NbMsg[i] = numEmails
            locstr = Addon.getLocalizedString(id=610) #messages(s)
            #Si il y a des msgs sur le serveur
            #if numEmails != 0:
                #Si affichage alternatif (les serveurs les uns apres les autres)
                #On affiche/stocke un seul serveur
            if ((ALT.lower() == 'true') and (i == NoServ)):
                msg = "%s : %d " % (NOM, numEmails) + "\n"
                #elif (ALT.lower() == 'false'):
                #Sinon on stocke le resultats de tout les serveurs
            else:
                msg = msg + "%s : %d " % (NOM, numEmails) + "\n"
                #Property pour afficher directement dans le skin avec le Home.xml
                #homeWin.setProperty( "server" , ("%s" % SERVER ))
            homeWin.setProperty( ("notifier.name%i" % i) , ("%s" % NOM ))
            homeWin.setProperty( ("notifier.msg%i" % i) , ("%i" % numEmails ))
                #debug( "name = %s %i" % (NOM, i))
                #debug( "numEmails = %i, Server : %i" % (numEmails, i))
            debug( "235 notifier.msg%i, Server : %s, numEmails : %i" % (i, NOM, numEmails))
            debug( "Affiche 236 : %s" % msg)
            numEmails = 0
            if NxMsgTot > 0:
                locstr = Addon.getLocalizedString(id=611) #Nouveau(x) message(s)
                #Affiche un popup a l'arrivee d'un nx msg
                if Addon.getSetting( 'popup' ) == "true":
                    xbmc.executebuiltin("XBMC.Notification( ,%d %s sur %s,160)" % (NxMsgTot, locstr, NOM))
    NoServ += 1  #On passe au serveur suivant
    if (NoServ > 3): NoServ = 1 #Si dernier serveur on revient au premier
    #On affiche soit directement dans le home, soit en utilsant le skin (SKIN = True)
    if (SKIN == "false"):
        MsgBox.setLabel( msg )
    else :
        MsgBox.setLabel( '' )

    #initialise start time
    start_time = time.time()
    time.sleep( .5 )
