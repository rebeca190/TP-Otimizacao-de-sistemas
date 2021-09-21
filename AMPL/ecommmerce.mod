set CIDADES;

param Custo {CIDADES, CIDADES} >= 0;
var X {CIDADES, CIDADES} binary;

minimize Caminho_Minimo:
sum {i in CIDADES, j in CIDADES} Custo[i, j] * X[i, j];

subject to Saidas_Vertice {i in CIDADES}:
sum {j in CIDADES} X[i,j] = 1;

subject to Chegada_Vertice {j in CIDADES}:
sum {i in CIDADES} X[i,j] = 1;

#Não escolhar Cidade 1 indo para Cidade 1
subject to Rest1{i in CIDADES}: sum {j in CIDADES: i = j} X[i,j] = 0;

#Não criar ciclos
subject to Rest2{a in CIDADES, b in CIDADES}: #Combinação com 2 cidades
(X[a,b] + X[b,a]) <= 1;

subject to Rest3{a in CIDADES, b in CIDADES, c in CIDADES}: #Combinação com 3 cidades
(X[a,b] + X[b,a] + X[a,c] + X[c,a] + X[b,c] + X[c,b] ) <= 2;


#subject to Rest4{a in CIDADES, b in CIDADES, c in CIDADES, d in CIDADES}: #Combinação com 4 cidades
#(X[a,b] + X[a,c] + X[a,d] + X[b,a] + X[b,c] + X[b,d] + X[c,a] + X[c,b] + X[c,d] + X[d,a] + X[d,b] + X[d,c]) <= 3;

#subject to rest5{a in CIDADES, b in CIDADES, c in CIDADES, d in CIDADES, e in CIDADES}:  #Combinação com 5 cidades
#(X[a,b] + X[b,a] + X[a,c] + X[c,a] + X[a,d] + X[d,a] + X[a,e] + X[e,a] + X[b,c] + X[c,b] + X[b,d] + X[d,b] + X[b,e] + X[e,b] + X[c,d] + X[d,c] + X[c,e] + X[e,c] + X[d,e] + X[e,d]) <= 4;

#subject to rest6{a in CIDADES, b in CIDADES, c in CIDADES, d in CIDADES, e in CIDADES, f in CIDADES}: #Combinação com 6 cidades
#(X[a,b] + X[a,c] + X[a,d] + X[a,e] + X[a,f] + X[b,a] + X[b,c] + X[b,d] + X[b,e] + X[b,f] + X[c,a] + X[c,b] + X[c,d] + X[c,e] + X[c,f] + X[d,a] + X[d,b] + X[d,c] + X[d,e] + X[d,f] + X[e,a] + X[e,b] + X[e,c] + X[e,d] + X[e,f] + X[f,a] + X[f,b] + X[f,c] + X[f,d] + X[f,e]) <= 5