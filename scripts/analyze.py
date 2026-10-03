import pymupdf,sys
NAVY=(0.1255,0.2275,0.3647)
def near(a,b,t=0.02): return all(abs(x-y)<t for x,y in zip(a,b))
def tables(path):
    d=pymupdf.open(path); out=[]
    for pi,p in enumerate(d):
        dr=p.get_drawings()
        hdr=[x['rect'] for x in dr if x.get('fill') and near(x['fill'],NAVY) and x['rect'].height>20 and x['rect'].width>15 and x['rect'].y0>60 and x['rect'].height<45]
        rowc=[x['rect'] for x in dr if x['rect'].width>20 and x['rect'].width<26 and 15<x['rect'].height<45 and abs(x['rect'].x0-31)<2 and not (x.get('fill') and near(x['fill'],NAVY))]
        ys=sorted(set(round(r.y0,1) for r in hdr))
        for y in ys:
            cols=sorted(set((round(r.x0,1),round(r.x1,1)) for r in hdr if abs(r.y0-y)<1))
            hh=[r for r in hdr if abs(r.y0-y)<1][0].height
            rows=sorted((r.y0,r.y1) for r in rowc if r.y0>y+hh-2)
            rs=[]; last=y+hh
            for a,b in rows:
                if abs(a-last)<3: rs.append((round(a,1),round(b,1))); last=b
                elif rs: break
            out.append((pi,y,hh,cols,rs))
    return out
if __name__=='__main__':
    for t in tables(sys.argv[1]): print('page',t[0]+1,'hdr_y',t[1],t[2],'cols',t[3],'rows',len(t[4]),t[4][:1],t[4][-1:])
