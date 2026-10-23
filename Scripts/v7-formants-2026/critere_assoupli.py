#!/usr/bin/env python3
"""Critere assoupli pour petits corpus. Seuil de residu (+1 dB) et test de
robustesse de fenetre INCHANGES. Deux leviers seulement :

 1. FUSION DES GROUPES DE RANG : 4 groupes (k=2-3,4-7,8-15,>=16) -> 2 (k=2-5, k>=6),
    et tolerance de bande portee de +/-1/12 a +/-1/6 d'octave. Meme test logique
    (le residu survit-il a un changement de rang harmonique ?), environ deux fois
    plus de notes distinctes par groupe.

 2. BOOTSTRAP PAR NOTE, GROUPE PAR GROUPE : au lieu d'un seuil de comptage arbitraire
    (">= 5 notes"), on reechantillonne les NOTES avec remise (les harmoniques d'une
    meme note sont correles : le bootstrap doit porter sur la note, pas sur le point)
    et on exige que la borne basse de l'intervalle de confiance a 90 % du residu
    median du groupe soit > 0. Un petit nombre de notes donne un intervalle large,
    donc un echec : l'incertitude est rendue honnetement.

 ABANDONNE : un test de permutation global sur la bande. Il donnait p<0.005 a presque
 tout, y compris aux candidats justement rejetes, parce que la bande a ete DETECTEE
 sur la positivite de son residu median puis testee sur cette meme positivite
 (biais de selection), et parce qu'il ne teste pas l'independance au rang.
"""
import os, sys, math, json
import numpy as np
rng = np.random.default_rng(20261023)
src = open(os.path.expanduser("~/work/analyse.py")).read().split("NAMES=[")[0]
ns={}; exec(compile(src,"c","exec"),ns)
load,harm,lres,comp,find = ns['load'],ns['harm'],ns['lres'],ns['comp'],ns['find']

KG_STRICT=[(2,3),(4,7),(8,15),(16,99)]
KG_FUSION=[(2,5),(6,99)]
NBOOT=3000
MIN_NOTES_GROUP=3     # plancher de testabilite d'un groupe
MIN_PTS_GROUP=5

def per_note(S,hw=0.75):
    out=[]
    for f0,mi,dy,sp in S:
        fr,am=harm(sp,f0)
        if fr.size==0: continue
        r,nl=lres(fr,am,hw); keep=nl>=3
        if keep.sum()==0: continue
        out.append((fr[keep],r[keep],np.arange(1,fr.size+1)[keep],mi))
    return out

def gstats_strict(notes,fc,tol=1/12):
    o=[]
    for lo,hi in KG_STRICT:
        v=[];ms=set()
        for fr,r,ks,m in notes:
            s=(ks>=lo)&(ks<=hi)&(np.abs(np.log2(fr/fc))<=tol)
            if s.any(): v.extend(r[s]); ms.add(m)
        o.append((len(v),len(ms),float(np.median(v)) if len(v)>=5 else None))
    return o

def split_rank(notes,fc,tol=1/6):
    """Rang median des points de la bande : garantit deux groupes peuples.
    La fusion a seuil fixe (k=2-5 / k>=6) perd le test quand tout le signal est
    a rang eleve (instruments tres graves) : le decoupage adaptatif l'evite."""
    ks=[k for fr,r,kk,m in notes for f,k in zip(fr,kk) if abs(math.log2(f/fc))<=tol]
    return int(np.median(ks)) if ks else None

def group_boot(notes,fc,lo,hi,tol=1/6):
    """Bootstrap par note du residu median du groupe de rang."""
    per=[]
    for fr,r,ks,m in notes:
        s=(ks>=lo)&(ks<=hi)&(np.abs(np.log2(fr/fc))<=tol)
        if s.any(): per.append(r[s])
    npts=sum(len(x) for x in per)
    if len(per)<MIN_NOTES_GROUP or npts<MIN_PTS_GROUP:
        return dict(testable=False,n_notes=len(per),n_pts=npts)
    obs=float(np.median(np.concatenate(per)))
    idx=rng.integers(0,len(per),size=(NBOOT,len(per)))
    meds=np.empty(NBOOT)
    for b in range(NBOOT):
        meds[b]=np.median(np.concatenate([per[i] for i in idx[b]]))
    lo5=float(np.percentile(meds,5))
    return dict(testable=True,n_notes=len(per),n_pts=npts,median=round(obs,2),
                ci_low=round(lo5,2),passe=bool(obs>=1.0 and lo5>0))

ALL={"bois":[("Petite flute","Piccolo"),("Flute","Flute"),("Flute basse","Bass-Flute"),
("Flute contrebasse","Contrabass-Flute"),("Hautbois","Oboe"),("Cor anglais","EnglishHorn"),
("Basson","Bassoon"),("Contrebasson","Contrebasson"),("Clarinette en Mib","Clarinet-Eb"),
("Clarinette en Sib","Clarinet_Bb"),("Clarinette basse","Bass-Clarinet-Bb"),
("Clarinette contrebasse","Contrabass-Clarinet-Bb")],
"cuivres":[("Cor","Horn"),("Trompette","Trumpet_C"),("Trombone tenor","Trombone"),
("Trombone basse","Bass-Trombone"),("Tuba basse","Bass_Tuba"),("Tuba contrebasse","Contrabass-Tuba")],
"cordes":[("Violon","Violin"),("Alto","Viola"),("Violoncelle","Violoncello"),("Contrebasse","Contrabass")]}

fam=sys.argv[1]; O=os.path.expanduser("~/work/out/"); REP={}
for disp,f in ALL[fam]:
    S=load(O+f+".txt")
    med,nn,npt=comp(S); F=find(med,nn,npt)
    rob={}
    for hw in (0.5,1.0):
        m2,n2,p2=comp(S,hw); rob[hw]=[a for a,b,c,d,e in find(m2,n2,p2)]
    notes=per_note(S); ndist=len({m for f0,m,dy,sp in S})
    rows=[]
    for fc,amp,blo,bhi,nnd in F:
        gs=gstats_strict(notes,fc)
        ok_s=sum(1 for npts,nms,mv in gs[1:] if mv is not None and mv>=1.0 and nms>=5)
        ok_s+= 1 if (gs[0][2] is not None and gs[0][2]>=1.0 and gs[0][1]>=5) else 0
        # recompte strict sur les groupes k>=2 uniquement (comme 00-METHODE)
        ok_s=sum(1 for npts,nms,mv in gs if mv is not None and mv>=1.0 and nms>=5)
        rb=all(any(abs(math.log2(p/fc))<=1/3 for p in rob[hw]) for hw in (0.5,1.0))
        vs="RETENU" if (ok_s>=2 and rb) else ("fragile" if ok_s>=2 or rb else "REJETE")
        kmed=split_rank(notes,fc)
        KGA=[(1,kmed),(kmed+1,99)] if kmed else KG_FUSION
        gb=[group_boot(notes,fc,lo,hi) for lo,hi in KGA]
        ntest=sum(1 for g in gb if g['testable'])
        npass=sum(1 for g in gb if g.get('passe'))
        if ntest<2:
            vr="non mesurable"
        elif npass==2 and rb: vr="RETENU-cr"
        elif (npass==2) != rb: vr="fragile-cr"   # un seul des deux tests passe
        else: vr="REJETE"
        rows.append(dict(fc=round(fc,1),amp=round(amp,2),strict=vs,relax=vr,rob=rb,
                         ok_s=ok_s,groupes=gb,k_median=kmed,
                         kgroupes=[f"<={kmed}",f">{kmed}"] if kmed else ["2-5",">=6"]))
    REP[disp]=dict(n=len(S),ndist=ndist,rows=rows)
    print(f"{disp:24s} n={len(S):4d} notes={ndist:3d}")
    for r in rows:
        lab=r.get('kgroupes',[("?","?")])
        g=" | ".join(("k%s: n/a(%dn)"%(lab[i],x['n_notes']) if not x['testable']
                      else "k%s: %+.1f [IC %+.1f] %dn"%(lab[i],x['median'],x['ci_low'],x['n_notes']))
                     for i,x in enumerate(r['groupes']))
        print(f"   {r['fc']:7.0f}Hz {r['strict']:8s} -> {r['relax']:14s}  {g}")
json.dump(REP,open(os.path.expanduser(f"~/work/relache_{fam}.json"),"w"))
