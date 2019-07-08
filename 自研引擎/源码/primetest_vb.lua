
N=vb.inputbox("Input max asadawdnumber:");

N=tonumber(N);

if N==nil or N<2 then
  msgbox("Wrong input\n",48);
 
 vb.msgbox("Wrong input\n",48);
 
 return;
end
T=os.clock();

pi_n=0;
if N>2 then pi_n=1;
s="2"; 
end;

for n=3,N,2 do
  I=n^0.5+1
  t=true
  for i=3,I,2 do
    if n%i==0 then
      t=false
      break
    end
  end
  if t then
    s=s.." "..n;
    pi_n=pi_n+1;

  end
end
--io.write("\nPrime Count="..pi_n..",Time="..(os.clock()-T).."\n");
vb.msgbox(s.."\nPrime Count="..pi_n..",Time="..(os.clock()-T));
