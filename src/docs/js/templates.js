// Modelos de documento.
const year = new Date().getFullYear();
const today = new Date().toLocaleDateString('pt-BR', { day: 'numeric', month: 'long', year: 'numeric' });

export const TEMPLATES = {
  blank: { title: 'Documento sem título', preset: 'padrao', html: '<p><br></p>' },

  abnt: {
    title: 'Trabalho acadêmico',
    preset: 'abnt',
    margins: { top: 3, right: 2, bottom: 2, left: 3 },
    html: `
<p class="cover" style="text-align:center"><strong>UNIVERSIDADE / INSTITUIÇÃO</strong></p>
<p class="cover" style="text-align:center">CURSO</p>
<p class="cover" style="text-align:center"><br></p><p class="cover" style="text-align:center"><br></p>
<p class="cover" style="text-align:center">NOME DO AUTOR</p>
<p class="cover" style="text-align:center"><br></p><p class="cover" style="text-align:center"><br></p><p class="cover" style="text-align:center"><br></p>
<p class="cover" style="text-align:center"><strong>TÍTULO DO TRABALHO</strong></p>
<p class="cover" style="text-align:center">Subtítulo (se houver)</p>
<p class="cover" style="text-align:center"><br></p><p class="cover" style="text-align:center"><br></p><p class="cover" style="text-align:center"><br></p><p class="cover" style="text-align:center"><br></p>
<p class="cover" style="text-align:center">CIDADE</p>
<p class="cover" style="text-align:center">${year}</p>
<div class="page-break"></div>
<div class="toc" contenteditable="false"><div class="toc-title">Sumário</div></div>
<div class="page-break"></div>
<h1>1 Introdução</h1>
<p>Apresente o tema, o problema de pesquisa, os objetivos e a justificativa do trabalho. O texto segue as normas da ABNT: fonte de 12 pt, espaçamento de 1,5 entre linhas, recuo de 1,25 cm na primeira linha e texto justificado.</p>
<h2>1.1 Objetivos</h2>
<p>Descreva o objetivo geral e os objetivos específicos.</p>
<h1>2 Desenvolvimento</h1>
<p>Desenvolva a fundamentação teórica e a metodologia. Citações diretas com mais de três linhas devem ter recuo de 4 cm, fonte menor e espaçamento simples.</p>
<h2>2.1 Metodologia</h2>
<p>Explique como a pesquisa foi conduzida.</p>
<h1>3 Conclusão</h1>
<p>Retome os objetivos e apresente as considerações finais.</p>
<h1 style="text-align:center">Referências</h1>
<p style="text-indent:0;text-align:left;line-height:1">SOBRENOME, Nome. <strong>Título da obra</strong>: subtítulo. Edição. Cidade: Editora, ano.</p>`,
  },

  report: {
    title: 'Relatório',
    preset: 'moderno',
    html: `
<p class="title">Relatório trimestral</p>
<p class="subtitle">Resultados e próximos passos · ${today}</p>
<div class="toc" contenteditable="false"><div class="toc-title">Sumário</div></div>
<h1>Resumo executivo</h1>
<p>Um parágrafo curto com as principais conclusões do período. Destaque os números mais importantes e o que eles significam para o negócio.</p>
<div class="callout"><span class="callout-icon">💡</span><div class="callout-body"><strong>Destaque:</strong> a receita cresceu 18% em relação ao trimestre anterior.</div></div>
<h1>Indicadores</h1>
<table><tbody>
<tr><th>Indicador</th><th>Meta</th><th>Resultado</th><th>Status</th></tr>
<tr><td>Receita</td><td>R$ 1,2 mi</td><td>R$ 1,35 mi</td><td>✅ Acima</td></tr>
<tr><td>Novos clientes</td><td>300</td><td>284</td><td>⚠️ Próximo</td></tr>
<tr><td>Satisfação (NPS)</td><td>60</td><td>67</td><td>✅ Acima</td></tr>
</tbody></table>
<h1>Análise</h1>
<h2>O que funcionou</h2>
<ul><li>Campanha de lançamento com conversão acima do esperado.</li><li>Redução do tempo de atendimento em 22%.</li></ul>
<h2>Pontos de atenção</h2>
<ul><li>Custos de aquisição em alta nos canais pagos.</li></ul>
<h1>Próximos passos</h1>
<ul class="checklist"><li data-checked="false">Revisar orçamento de marketing</li><li data-checked="false">Lançar programa de indicação</li><li data-checked="true">Contratar analista de dados</li></ul>`,
  },

  letter: {
    title: 'Carta',
    preset: 'classico',
    html: `
<p style="text-align:right;text-indent:0">São Paulo, ${today}.</p>
<p style="text-indent:0"><br></p>
<p style="text-indent:0">À<br>Empresa Destinatária<br>A/C: Nome do Responsável</p>
<p style="text-indent:0"><br></p>
<p style="text-indent:0"><strong>Assunto:</strong> Escreva aqui o assunto da carta</p>
<p style="text-indent:0">Prezado(a) Senhor(a),</p>
<p>Escreva o primeiro parágrafo apresentando o motivo do contato de forma clara e objetiva.</p>
<p>Desenvolva os detalhes necessários no segundo parágrafo, incluindo datas, valores ou referências importantes.</p>
<p>Agradeço a atenção e coloco-me à disposição para quaisquer esclarecimentos.</p>
<p style="text-indent:0">Atenciosamente,</p>
<p style="text-indent:0"><br></p>
<p style="text-indent:0"><strong>Seu Nome</strong><br>Cargo · Empresa<br>(11) 90000-0000 · email@exemplo.com</p>`,
  },

  resume: {
    title: 'Currículo',
    preset: 'moderno',
    margins: { top: 1.8, right: 2, bottom: 1.8, left: 2 },
    html: `
<p class="title">Seu Nome Completo</p>
<p class="subtitle">Cargo pretendido · Cidade, UF · (11) 90000-0000 · email@exemplo.com · linkedin.com/in/seunome</p>
<hr>
<h2>Resumo profissional</h2>
<p>Profissional com X anos de experiência em ..., com foco em resultados como ... Busco oportunidade como ... para contribuir com ...</p>
<h2>Experiência</h2>
<h3>Cargo — Empresa</h3>
<p style="color:#5f6b7a"><em>Jan 2022 – atual · Cidade</em></p>
<ul><li>Principal conquista com número: aumentou X em Y%.</li><li>Responsável por ..., liderando equipe de N pessoas.</li></ul>
<h3>Cargo anterior — Empresa</h3>
<p style="color:#5f6b7a"><em>Mar 2019 – Dez 2021 · Cidade</em></p>
<ul><li>Conquista relevante e mensurável.</li></ul>
<h2>Formação</h2>
<p><strong>Curso</strong> — Instituição (ano de conclusão)</p>
<h2>Competências</h2>
<p>Competência 1 · Competência 2 · Competência 3 · Inglês avançado · Pacote Office</p>`,
  },

  meeting: {
    title: 'Ata de reunião',
    preset: 'padrao',
    html: `
<p class="title">Ata de reunião</p>
<table><tbody>
<tr><th>Data</th><td>${today}</td><th>Horário</th><td>10h00 – 11h00</td></tr>
<tr><th>Local</th><td>Sala 3 / Online</td><th>Responsável</th><td>Nome</td></tr>
</tbody></table>
<h2>Participantes</h2>
<ul><li>Nome — Área</li><li>Nome — Área</li></ul>
<h2>Pauta</h2>
<ol><li>Tema 1</li><li>Tema 2</li><li>Assuntos gerais</li></ol>
<h2>Discussões</h2>
<p>Resumo do que foi discutido em cada item da pauta.</p>
<h2>Decisões</h2>
<div class="callout success"><span class="callout-icon">✅</span><div class="callout-body">Decisão tomada e aprovada pelo grupo.</div></div>
<h2>Ações</h2>
<table><tbody>
<tr><th>Ação</th><th>Responsável</th><th>Prazo</th></tr>
<tr><td>Descrever a ação</td><td>Nome</td><td>dd/mm</td></tr>
</tbody></table>`,
  },
};
