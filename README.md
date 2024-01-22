# Sommaire de la documentation

## 1. Principales informations

La description du logiciel se trouve dans la [Notes de conception](./docs/Conception.md).

## 2. Gestion de l'environnement virtuel

Créer et activer un nouvel environnement virtuel avec les commandes:

```python
python3 -m venv .venv
source .venv/bin/activate
```

Vous pouvez sortir de cet environnement virtuel avec la commande:

```python
deactivate
```

## 3. Gestion des dépendances

Il est possible d'ajouter une nouvelle dépendance au fichier requirement avec la commande:

```python
pip freeze > requirements.txt
```

Pour installer les dépendances depuis le fichier `requirements.txt`, utiliser la commande:

```python
pip install -r requirements.txt
```
