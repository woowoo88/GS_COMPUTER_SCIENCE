# Mission Control AI

Simulação interativa de um sistema de alerta por lógica digital para uma missão espacial experimental.

O sistema monitora seis condições operacionais e aciona a saída `X` quando identifica uma combinação crítica.

## Executar

Abra [`simulacao_mission_control.html`](simulacao_mission_control.html) em um navegador.

Na simulação é possível:

- alternar as seis entradas digitais;
- acompanhar os sinais intermediários;
- observar o LED virtual de alerta;
- consultar as 64 combinações da tabela verdade.

## Expressão simplificada

```text
X = AC + EF + D(B + NOT(A))
```

O sistema utiliza as operações lógicas `AND`, `OR` e `NOT`.

## Informações acadêmicas

- **Curso:** Ciência da Computação
- **Turma:** 1CCPF
- **Professor:** Mauricio Neto
- **Data:** 08/06/2026
