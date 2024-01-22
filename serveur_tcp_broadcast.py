import socket
import threading
import time
import random

# Liste pour stocker les connexions clients
clients = []
clients_lock = threading.Lock()  # Verrou pour synchroniser l'accès à la liste

# Fonction pour gérer la communication avec un client individuel
def handle_client(client_socket, client_address):
    print(f"Connexion acceptée depuis {client_address}")

    # Ajouter la connexion client à la liste
    with clients_lock:
        clients.append(client_socket)

    while True:
        # Générer une valeur réelle aléatoire entre 0 et 200.0
        random_value = random.uniform(0, 200.0)

        # Préparer la trame à envoyer à tous les clients
        message = f"ROTATION=91.000WELDSPEED={random_value}INTERPASS_TEMP=25.300"

        # Envoyer la trame à tous les clients
        with clients_lock:
            for client in clients:
                try:
                    #client.send(message.encode('utf-8'))
                    client.send(message.encode('ascii'))
                except socket.error:
                    # En cas d'erreur, le client a été déconnecté, on le retire de la liste
                    clients.remove(client)

        # Attendre pendant 100 millisecondes avant d'envoyer la prochaine trame
        #time.sleep(0.1)
        # Attendre pendant 0.5 millisecondes (2000 Hz) avant d'envoyer la prochaine trame
        time.sleep(0.0005)


# Configurer le serveur
server_host = '127.0.0.1'
server_port = 12345

server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
server.bind((server_host, server_port))
server.listen(5)

print(f"Serveur en écoute sur {server_host}:{server_port}")

# Attendre les connexions des clients et gérer chaque client dans un thread séparé
while True:
    client, addr = server.accept()

    # Créer un thread pour gérer la communication avec le client
    client_handler = threading.Thread(target=handle_client, args=(client, addr))
    client_handler.start()
