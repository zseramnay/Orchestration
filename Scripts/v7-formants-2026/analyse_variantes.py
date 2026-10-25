#!/usr/bin/env python3
"""Analyse complete (descripteurs + verdicts strict et assoupli) sur une liste de configs."""
import os,sys,math,json
import numpy as np
rng=np.random.default_rng(20261023)
base=open(os.path.expanduser("~/work/analyse.py")).read().split("NAMES=[")[0]
ns={}; exec(compile(base,"c","exec"),ns)
load,harm,lres,comp,find,te = ns['load'],ns['harm'],ns['lres'],ns['comp'],ns['find'],ns['te']
NBOOT=2000; MINN=3; MINP=5
def per_note(S,hw=0.75):
    o=[]
    for f0,mi,dy,sp in S:
        fr,am=harm(sp,f0)
        if fr.size==0: continue
        r,nl=lres(fr,am,hw); k=nl>=3
        if k.sum(): o.append((fr[k],r[k],np.arange(1,fr.size+1)[k],mi))
    return o
def gb(notes,fc,lo,hi,tol=1/6):
    per=[]
    for fr,r,ks,m in notes:
        s=(ks>=lo)&(ks<=hi)&(np.abs(np.log2(fr/fc))<=tol)
        if s.any(): per.append(r[s])
    np_=sum(len(x) for x in per)
    if len(per)<MINN or np_<MINP: return dict(t=False,n=len(per))
    obs=float(np.median(np.concatenate(per)))
    idx=rng.integers(0,len(per),size=(NBOOT,len(per)))
    med=np.array([np.median(np.concatenate([per[i] for i in idx[b]])) for b in range(NBOOT)])
    lo5=float(np.percentile(med,5))
    return dict(t=True,n=len(per),med=round(obs,2),ci=round(lo5,2),p=bool(obs>=1.0 and lo5>0))
def spear(a,b):
    ra=np.argsort(np.argsort(a)).astype(float); rb=np.argsort(np.argsort(b)).astype(float)
    ra-=ra.mean(); rb-=rb.mean(); d=math.sqrt((ra**2).sum()*(rb**2).sum())
    return float((ra*rb).sum()/d) if d else float('nan')
O=os.path.expanduser("~/work/out/"); REP={}
for arg in sys.argv[1:]:
    disp,f=arg.split("=")
    S=load(O+f+".txt")
    if len(S)<5: print(f"!! {disp} vide"); continue
    med,nn,npt=comp(S); F=find(med,nn,npt)
    rob={}
    for hw in (0.5,1.0):
        m2,n2,p2=comp(S,hw); rob[hw]=[a for a,b,c,d,e in find(m2,n2,p2)]
    notes=per_note(S); f0s=np.array([x[0] for x in S])
    tv=np.array([te(sp,f0)[0] for f0,mi,dy,sp in S],float); g=~np.isnan(tv)
    rho=spear(f0s[g],tv[g]) if g.sum()>5 else float('nan')
    sl=[];nh=[];od=[]
    for f0,mi,dy,sp in S:
        fr,am=harm(sp,f0); nh.append(fr.size)
        if fr.size>=5:
            x=np.log2(fr);A=np.vstack([x,np.ones_like(x)]).T
            sl.append(np.linalg.lstsq(A,am,rcond=None)[0][0])
        if fr.size>=9: od.append(np.mean([am[k] for k in(2,4,6,8)])-np.mean([am[k] for k in(1,3,5,7)]))
    rows=[]
    for fc,amp,blo,bhi,nnd in F:
        ks=[k for fr,r,kk,m in notes for ff,k in zip(fr,kk) if abs(math.log2(ff/fc))<=1/6]
        km=int(np.median(ks)) if ks else None
        G=[gb(notes,fc,1,km),gb(notes,fc,km+1,99)] if km else []
        nt=sum(1 for x in G if x['t']); npass=sum(1 for x in G if x.get('p'))
        rb=all(any(abs(math.log2(p/fc))<=1/3 for p in rob[hw]) for hw in (0.5,1.0))
        if nt<2: v="non mesurable"
        elif npass==2 and rb: v="retenu-cr"
        elif (npass==2)!=rb: v="fragile-cr"
        else: v="rejete"
        rows.append(dict(fc=round(fc,1),amp=round(amp,2),blo=round(blo,1),bhi=round(bhi,1),
                         verdict=v,rob=rb,k_med=km,groupes=G,notes=nnd))
    REP[disp]=dict(n=len(S),ndist=len({m for f0,m,dy,sp in S}),rho=round(rho,2),
        slope=round(float(np.median(sl)),1),slope_sd=round(float(np.std(sl)),1),
        nh=int(np.median(nh)),odd=(round(float(np.median(od)),1) if od else None),
        te_med=round(float(np.nanmedian(tv)),0),te_sd=round(float(np.nanstd(tv)),0),rows=rows)
    r=REP[disp]
    print(f"{disp:26s} n={r['n']:4d} rho={r['rho']:+.2f} pente={r['slope']:+6.1f} odd={('--' if r['odd'] is None else format(r['odd'],'+5.1f'))} | "
          +"  ".join(f"{x['fc']:.0f}:{x['verdict'][:9]}" for x in r['rows']))
json.dump(REP,open(os.path.expanduser(f"~/work/an2_{sys.argv[1].split('=')[0][:6]}.json"),"w"))
