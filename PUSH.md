# Push A and P so other people can clone

From a machine with git and your GitHub login:

```bash
unzip A-and-P.zip
cd A-and-P
git init
git add .
git commit -m "A and P: n = 1+0 catalog matrix"
git branch -M main
git remote add origin https://github.com/fitzyracing1/A-and-P.git
git push -u origin main
```

Create the empty public repo `fitzyracing1/A-and-P` on GitHub first if it does not exist.

Clone for everyone else:

```bash
git clone https://github.com/fitzyracing1/A-and-P.git
```
