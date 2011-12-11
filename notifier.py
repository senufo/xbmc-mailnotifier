import xbmc, xbmcgui
import xbmcaddon
import poplib, imaplib
import time

print "LIGNE 6"
__author__     = "Senufo"
__scriptid__   = "service.notifier"
__scriptname__ = "Notifier"

Addon          = xbmcaddon.Addon(__scriptid__)
print "LIGNE 12"

__cwd__        = Addon.getAddonInfo('path')
__version__    = Addon.getAddonInfo('version')
__language__   = Addon.getLocalizedString

__profile__    = xbmc.translatePath(Addon.getAddonInfo('profile'))
__resource__   = xbmc.translatePath(os.path.join(__cwd__, 'resources',
                                                 'lib'))

#sys.path.append (__resource__)
print "ligne 23"


#ID de la fenetre HOME
WINDOW_HOME = 10000
Home = 10000

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
ALT = Addon.getSetting('alt')
MsgBox = None
MsgBoxId = None
start_time = 0
re_added_control = False
print "ligne 48"
#Verifie que xbmc tourne
while (not xbmc.abortRequested):
    #Attente avant de relever les mails
    intervalle = int(float(Addon.getSetting('time')) * 60.0)
    if start_time and (time.time() - start_time) < intervalle:
        time.sleep(.5)
        SHOW_UPDATE     = Addon.getSetting('show_update') == "true"
        if MsgBox:
            # optional show le temps qu'il reste avant la prochaine MAJ
            try:
                if SHOW_UPDATE:
                    up = int(intervalle) - (time.time() - start_time)
                    locstr = Addon.getLocalizedString(615)  #Update in %i second
                    #print "MSG up = %s " % msg
                    label = "%s[CR]" %  msg + locstr % up
                else: #Il faut rafraichir l'affichage
                    #print "MSG = %s " % msg
                    label = '%s' % msg
                MsgBox.setLabel( label )
            except Exception, e:
                print str(e)

        HomeNotVisible = xbmc.getCondVisibility( "!Window.IsVisible(10000)" )
        if HomeNotVisible:
            #oop! on est plus sur le home
            re_added_control = True
        elif re_added_control and not HomeNotVisible:
            MsgBox = xbmcgui.ControlLabel( x, y, width, height, '', font, color )
            # add control label and set default label
            homeWin.addControl( MsgBox )
            # get control id
            MsgBoxId = MsgBox.getId()
            re_added_control = False
            # reload addon setting possible change
            Addon = xbmcaddon.Addon( __scriptid__ )

        # continue le while sans faire le reste
        continue

    homeWin = xbmcgui.Window(WINDOW_HOME)
    homeWin.setProperty( "username", "FROST" )
    homeWin.setProperty( "server" , "SERVER" )
    homeWin.setProperty( "name" , "NOM" )
    homeWin.setProperty( "msg" , "numeEmails" )

    #xbmc.executebuiltin( "SetProperty(username,'eeee',10000)" )
    if MsgBoxId:
        try: MsgBox = homeWin.getControl( MsgBoxId )
        except: MsgBoxId = None
    if MsgBoxId is None:
        MsgBox = xbmcgui.ControlLabel( x, y, width, height, '', font, color )
        #retire le control s'il exist # pas vraiment besoin le test a ete fait avec homeWin.getControl( MsgBoxId )
        #try: homeWin.removeControl( MsgBox )
        #except: pass #print "Le controle n'existe pas"
        # add control label and set default label
        homeWin.addControl( MsgBox )
        # get control id
        MsgBoxId = MsgBox.getId()
    locstr = Addon.getLocalizedString(616) #Mise a jour
    MsgBox.setLabel( locstr % ' ' )

    #On vide le message
    msg = ''

    #On recupere les parametres des trois serveurs
    for i in range( 1, 4 ): #[1,2,3]:
        if Addon.getSetting( 'enableserver%i' % i ) == "false":
            #print "Enableserver = %s, i = %d  " % (Addon.getSetting(
            #    'enableserver%i' % i), i)
            continue
        USER     = Addon.getSetting( 'user%i'   % i )
        NOM      = Addon.getSetting( 'name%i'   % i )
        SERVER   = Addon.getSetting( 'server%i' % i )
        PASSWORD = Addon.getSetting( 'pass%i'   % i )
        PORT     = Addon.getSetting( 'port%i'   % i )
        SSL      = Addon.getSetting( 'ssl%i'    % i )
        TYPE     = Addon.getSetting( 'type%i'   % i )
        FOLDER   = Addon.getSetting( 'folder%i' % i )

        #print "SERVER = %s, PORT = %s, USER = %s, password = %s, SSL = %s" % (SERVER, PORT, USER, PASSWORD, SSL)
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
                    #print "IMAP numEmails = %d " % numEmails

                #print "numEmails = %d " % numEmails
                locstr = Addon.getLocalizedString(610) #message(s)
                #msg = msg + "%s : %d %s" % (NOM,numEmails, locstr) + "\n"
                #numEmails = 0
                #xbmc.executebuiltin( "SetProperty(username,%s,Home)" % "FROST" )
            except:
                locstr = Addon.getLocalizedString(613) #erreur de connection
                if Addon.getSetting( 'erreur' ) == "true":
                    xbmc.executebuiltin("XBMC.Notification(%s : ,%s,30)" % (locstr, SERVER))
                print "Erreur de connection : %s" % SERVER
#Msg affiche sur le HOME
                msg = msg + "%s : %s\n" % (NOM, locstr)

            if numEmails > 0:
                MsgTot = True #Il y a des messages
            #On regarde si un nouveau mail est arrive
            if NbMsg[i] == 0:
                NbMsg[i] = numEmails
            NxMsg = numEmails - NbMsg[i]
            if NxMsg > 0:
                NbMsg[i] = NbMsg[i] + NxMsg
                NxMsgTot = NxMsgTot + NxMsg
            else:
                NbMsg[i] = numEmails
            locstr = Addon.getLocalizedString(id=610) #messages(s)
            if numEmails != 0:
                if ((ALT.lower() == 'true') and (i == NoServ)):
                    msg = "%s : %d " % (NOM, numEmails) + "\n"
                elif (ALT.lower() == 'false'):
                    msg = msg + "%s : %d " % (NOM, numEmails) + "\n"
                #homeWin.setProperty( "username", "FROST" )
                homeWin.setProperty( "server" , ("%s" % SERVER ))
                homeWin.setProperty( "name" , ("%s" % NOM ))
                homeWin.setProperty( "msg" , ("%s" % numEmails ))
            numEmails = 0
            if NxMsgTot > 0:
                locstr = Addon.getLocalizedString(id=611) #Nouveau(x) message(s)
                #Affiche un popup a l'arrivee d'un nx msg
                if Addon.getSetting( 'popup' ) == "true":
                    xbmc.executebuiltin("XBMC.Notification( ,%d %s sur %s,160)" % (NxMsgTot, locstr, NOM))
    NoServ += 1  #On passe au serveur suivant
    if (NoServ > 3): NoServ = 1 #Si dernier serveur on revient au premier
    MsgBox.setLabel( msg )

    #initialise start time
    start_time = time.time()
    time.sleep( .5 )
