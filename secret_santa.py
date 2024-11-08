import pandas as pd
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import sys

def load_participants(file_path):
    """
    Charge les participants depuis un fichier Excel.
    Si la colonne 'Catégorie' n'existe pas, elle est créée et remplie avec des valeurs None.

    :param file_path: Chemin du fichier Excel contenant les informations des participants.
    :return: DataFrame Pandas avec les informations des participants.
    """
    df = pd.read_excel(file_path)
    if 'Catégorie' not in df.columns:
        df['Catégorie'] = None

    return df

def filter_valid_pairs(participants):
    """
    Génère toutes les paires valides de donneur-receveur en respectant les règles:
    - Un participant ne peut pas tirer une personne de la même catégorie.
    - Un participant ne peut pas se tirer lui-même.

    :param participants: DataFrame Pandas avec les informations des participants.
    :return: Liste de tuples (index donneur, index receveur) pour les paires valides.
    """
    pairs = []
    for i, giver in participants.iterrows():
        for j, receiver in participants.iterrows():
            # Vérifie que le donneur et le receveur ne sont pas la même personne et qu'ils n'ont pas la même catégorie
            if (giver['NOM'] != receiver['NOM'] or giver['Prénom'] != receiver['Prénom']) and giver['Catégorie'] != receiver['Catégorie']:
                pairs.append((i, j))

    return pairs

def secret_santa_draw(participants):
    """
    Réalise le tirage au sort Secret Santa en attribuant à chaque participant un destinataire valide.
    Assure qu'il n'y a pas de participant tirant une personne de la même catégorie ou lui-même.

    :param participants: DataFrame Pandas avec les informations des participants.
    :return: Liste de tuples (donneur, receveur) où chaque élément est une ligne du DataFrame des participants.
    :raises Exception: si un tirage valide ne peut pas être généré.
    """
    pairs = filter_valid_pairs(participants)
    random.shuffle(pairs)
    matched = {}
    
    for giver, receiver in pairs:
        if giver not in matched and receiver not in matched.values():
            matched[giver] = receiver
        if len(matched) == len(participants):
            break
    
    if len(matched) != len(participants):
        raise Exception("Impossible de générer un tirage Secret Santa valide avec les contraintes actuelles.")
    
    results = [(participants.iloc[giver], participants.iloc[receiver]) for giver, receiver in matched.items()]
    return results

def save_results_to_excel(results):
    """
    Sauvegarde les résultats du tirage Secret Santa dans un fichier Excel avec un nom incluant la date et l'heure.

    :param results: Liste de tuples (donneur, receveur) contenant les résultats du tirage.
    :return: Chemin du fichier Excel de sortie.
    """
    now = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"secret_santa_results_{now}.xlsx"
    output_data = [{'NOM Donneur': row[0]['NOM'], 'Prénom Donneur': row[0]['Prénom'], 'NOM Destinataire': row[1]['NOM'], 'Prénom Destinataire': row[1]['Prénom']} for row in results]
    output_df = pd.DataFrame(output_data)
    output_df.to_excel(output_file, index=False)
    return output_file

def send_emails(results, smtp_server, smtp_port, email, password):
    """
    Envoie un email à chaque participant avec les informations de leur destinataire Secret Santa.

    :param results: Liste de tuples (donneur, receveur) contenant les résultats du tirage.
    :param smtp_server: Serveur SMTP pour envoyer les emails.
    :param smtp_port: Port du serveur SMTP.
    :param email: Adresse email de l'expéditeur.
    :param password: Mot de passe de l'adresse email de l'expéditeur.
    """
    for giver, receiver in results:
        # Crée le message email
        msg = MIMEMultipart()
        msg['From'] = email
        msg['To'] = giver['Email']
        msg['Subject'] = "Votre Secret Santa 🎅"
        body = f"Bonjour {giver['Prénom']},\n\nVous avez été désigné pour offrir un cadeau à {receiver['Prénom']} {receiver['NOM']} ! 🎄\n\nJoyeux Noël !"
        msg.attach(MIMEText(body, 'plain'))

        # Envoie l'email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(email, password)
            server.sendmail(email, giver['Email'], msg.as_string())

def main():
    """
    Point d'entrée principal du script.
    Vérifie les arguments de la ligne de commande pour déterminer le mode (-test ou -send).
    En mode -test, le tirage est enregistré dans un fichier Excel.
    En mode -send, le tirage est enregistré dans un fichier Excel et envoyé par email à chaque participant.
    """
    if len(sys.argv) < 3:
        print("Usage: python secret_santa.py <file_path> -test|-send")
        sys.exit(1)

    file_path = sys.argv[1]
    mode = sys.argv[2]
    
    # Charge les participants et effectue le tirage
    participants = load_participants(file_path)
    results = secret_santa_draw(participants)
    
    if mode == '-test':
        # Mode test : Sauvegarde uniquement le tirage dans un fichier Excel
        output_file = save_results_to_excel(results)
        print(f"Tirage enregistré dans {output_file} en mode test.")
    elif mode == '-send':
        # Mode envoi : Sauvegarde et envoie les emails
        output_file = save_results_to_excel(results)
        print(f"Tirage enregistré dans {output_file}. Envoi des emails en cours...")
        
        # Informations de connexion SMTP (à configurer)
        smtp_server = "smtp.example.com"
        smtp_port = 587
        email = "your_email@example.com"
        password = "your_password"
        
        send_emails(results, smtp_server, smtp_port, email, password)
        print("Emails envoyés.")
    else:
        print("Argument non reconnu. Utilisez '-test' pour tester ou '-send' pour envoyer les emails.")
        sys.exit(1)

if __name__ == "__main__":
    main()