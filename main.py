# Importer la bibliothèque PyWin32
import win32gui
import win32con

# Définir une fonction qui liste les programmes ouverts
def lister_programmes():
    # Créer une liste vide
    programmes = []
    # Définir une fonction qui ajoute le nom et le handle d'une fenêtre à la liste
    def callback(handle, data):
        # Récupérer le nom de la fenêtre
        nom = win32gui.GetWindowText(handle)
        # Si la fenêtre est visible et a un nom
        if win32gui.IsWindowVisible(handle) and nom:
            # Ajouter le nom et le handle à la liste
            programmes.append((nom, handle))
    # Parcourir toutes les fenêtres ouvertes et appeler la fonction callback
    win32gui.EnumWindows(callback, None)
    # Retourner la liste
    return programmes

# Définir une fonction qui affiche un menu pour choisir un programme
def choisir_programme():
    # Récupérer la liste des programmes
    programmes = lister_programmes()
    # Afficher un message
    print("Choisissez un programme qui restera en avant plan :")
    # Parcourir la liste des programmes
    for i, (nom, handle) in enumerate(programmes):
        # Afficher le numéro et le nom du programme
        print(f"{i+1} - {nom}")
    # Demander à l'utilisateur de saisir un numéro
    choix = int(input("Entrez le numéro du programme : "))
    # Vérifier que le numéro est valide
    if 1 <= choix <= len(programmes):
        # Récupérer le handle du programme choisi
        handle = programmes[choix-1][1]
        # Retourner le handle
        return handle
    else:
        # Afficher un message d'erreur
        print("Numéro invalide")
        # Retourner None
        return None

# Définir une fonction qui met un programme en avant plan
def mettre_en_avant_plan(handle):
    # Vérifier que le handle est valide
    if handle:
        # Récupérer le handle de la fenêtre active
        active = win32gui.GetForegroundWindow()
        # Si le handle est différent de la fenêtre active
        if handle != active:
            # Récupérer le thread de la fenêtre active
            thread_active = win32gui.GetWindowThreadProcessId(active)[0]
            # Récupérer le thread du programme
            thread_programme = win32gui.GetWindowThreadProcessId(handle)[0]
            # Attacher les threads
            win32gui.AttachThreadInput(thread_active, thread_programme, True)
            # Mettre le programme en avant plan
            win32gui.SetWindowPos(handle, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            # Détacher les threads
            win32gui.AttachThreadInput(thread_active, thread_programme, False)

# Appeler la fonction choisir_programme et récupérer le handle
handle = choisir_programme()
# Appeler la fonction mettre_en_avant_plan avec le handle
mettre_en_avant_plan(handle)
