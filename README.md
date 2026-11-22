# Instalare WordPress

## Cerințe și acces

- **Versiunea minimă PHP:** 8.3
- **cPanel:** sunt necesare datele de autentificare (utilizator și parolă).
- **WordPress Admin:** păstrați datele de autentificare (utilizator și parolă) într-un loc sigur.

## Instalarea fișierelor pe o instalare existenta WordPress

1. Încărcați `wp-content.zip` pe server.
2. Dezarhivați-l și permiteți înlocuirea fișierelor existente.

## Configurarea WordPress

- Activați pluginurile necesare= SetareDate, SetareDrive, Postie, PDF Embedder
- Configurați biblioteca media să nu creeze foldere pe lună/an.
- Adăugați cel putin categoria **articole**.

## Google Drive

- Spațiu recomandat: **1.500 MB**; de preferat să existe deja un cont Google Drive.
- **Folderul Monitorului Oficial:** ID-ul folderului trebuie completat pentru locația Drive asociată domeniului `primariabebaveche.ro`.
- **API Google Drive:** configurarea se face în [Google Cloud Console](https://console.cloud.google.com/).
- Pentru salvarea directă pe server, folosiți locația `/wp-content/upload/formulare` (confirmați calea exactă înainte de configurare).

## Publicarea articolelor prin e-mail (Postie)

Configurați pluginul **Postie** cu IMAP prin SSL, portul **993**, fără corectare de timp.

### Variante de e-mail

- **Gmail:** `imap.gmail.com` — notează utilizatorul `brwb8w5a@gmail.com`; parola trebuie introdusă separat și păstrată securizat.
- **E-mail intern recomandat:** `wp@primariabebaveche.ro`, server `mail.primariabebaveche.ro`; parola trebuie introdusă separat și păstrată securizat.

### Setări de publicare

- Configurați lista adreselor de e-mail autorizate să publice. Administratorul site-ului WordPress este deja autorizat.
- E-mailurile cu titlul care începe cu `[public]` sunt publicate direct pe prima pagină.
- **Utilizator implicit pentru postări:** `admin`
- **Categorie implicită:** `articole`

## Sincronizarea fișierelor

Au fost adaugate deja pe desktop scurtaturi pentru programul de email automat, pentru adaugare prin drag&drop de fisiere
si programul de sincronizare a fisierelor noi pe server
trebuie doar configurate setarile - nume server (fara http sau ftp) user, parola, adresa de email setata in Postie si parola emailului.
Atat!
