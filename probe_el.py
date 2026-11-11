import json, sys, time, urllib.request
URL="https://prozorro.gov.ua/api/search/tenders?page=1"
def q(body):
    for i in range(5):
        try:
            r=urllib.request.Request(URL,json.dumps(body).encode(),{"Content-Type":"application/json","User-Agent":"research"})
            time.sleep(1.1); return json.load(urllib.request.urlopen(r,timeout=60))
        except Exception as e: print("err",e); time.sleep(5)
    return {}
M={"start":"2026-03-01","end":"2026-03-31"}
PT=["aboveThreshold","belowThreshold","priceQuotation","reporting","negotiation","negotiation.quick","aboveThresholdUA","aboveThresholdEU","simple.defense","closeFrameworkAgreementUA","competitiveOrdering"]
for name,base in [("electricity 09310000-5",{"cpv":["09310000-5"]}),("fuel",{"cpv":["09130000-9","09132000-3","09134000-7","09134200-9","09133000-0","09122000-0","09122110-4","09132100-4","09134220-5","30163100-0"]})]:
    print("==",name,"all:",q({**base,"date":{"notice":M}}).get("total"))
    for p in PT: print("  ",p,q({**base,"proc_type":[p],"date":{"notice":M}}).get("total"))
    for y in [("2025-01-01","2025-12-31"),("2026-01-01","2026-11-11")]:
        print("  year",y[0][:4],q({**base,"date":{"notice":{"start":y[0],"end":y[1]}}}).get("total"))
r=q({"cpv":["09310000-5"],"proc_type":["reporting"],"date":{"notice":M}})
print(json.dumps((r.get("data") or [None])[0],ensure_ascii=False)[:1500])
r=q({"cpv":["09310000-5"],"date":{"notice":M}})
print(json.dumps((r.get("data") or [None])[0],ensure_ascii=False)[:1500])
