Title: git
Date: 2019-01-02 00:22
Author: Lulef
Category: Sammlung
Slug: git
Status: published
```
git init
git remote add origin git@git.sr.ht:~chincherpa/ascii_print
git remote add origin git@git.sr.ht:~chincherpa/my_scripts
git add *
git commit -m "initial commit"
git push origin master
```
#### Datei löschen
```
git rm file1.txt
git commit -m "remove file1.txt" 
```
But if you want to remove the file only from the Git repository and not remove it from the filesystem, use:
```
git rm --cached file1.txt
git commit -m "remove file1.txt"
```
And to push changes to remote repo
```
git push origin branch_name
```
