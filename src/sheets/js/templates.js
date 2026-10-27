// Modelos de planilha prontos.
import { Workbook } from './model.js';
import { parseAddr, serialFromParts } from './values.js';

const HEAD = { b: true, color: '#ffffff', bg: '#2f9e6e', ha: 'center', va: 'middle' };
const TITLE = { b: true, size: 18, color: '#1d2129' };
const SUB = { color: '#6b7280', i: true };
const CUR = { fmt: 'currency', dec: 2 };
const thin = (color = '#d0d7de') => ({ t: { w: 'thin', color }, r: { w: 'thin', color }, b: { w: 'thin', color }, l: { w: 'thin', color } });

function put(s, a, v, style) {
  const { r, c } = parseAddr(a);
  if (typeof v === 'string' && v.startsWith('=')) s.set(r, c, { v: null, f: v.slice(1), s: style });
  else s.set(r, c, { v, s: style });
}

function row(s, r, values, style) {
  values.forEach((v, i) => put(s, `${String.fromCharCode(65 + i)}${r}`, v, typeof style === 'function' ? style(i) : style));
}

export const SHEET_TEMPLATES = {
  blank: { title: 'Planilha sem título', build: () => new Workbook() },

  budget: {
    title: 'Orçamento pessoal',
    build() {
      const wb = new Workbook();
      const s = wb.sheet;
      s.name = 'Orçamento';
      put(s, 'A1', 'Orçamento pessoal', TITLE);
      put(s, 'A2', 'Preencha os valores previstos e reais. Os totais e o gráfico se atualizam sozinhos.', SUB);
      row(s, 4, ['Categoria', 'Tipo', 'Previsto', 'Real', 'Diferença', '% do previsto'], HEAD);
      const items = [
        ['Salário', 'Receita', 6500, 6500], ['Freelas', 'Receita', 800, 1150],
        ['Aluguel', 'Despesa', 1800, 1800], ['Mercado', 'Despesa', 1100, 1243.9], ['Contas (luz, água, internet)', 'Despesa', 420, 398.5],
        ['Transporte', 'Despesa', 450, 512], ['Saúde', 'Despesa', 380, 380], ['Lazer', 'Despesa', 500, 655], ['Educação', 'Despesa', 300, 300],
        ['Investimentos', 'Despesa', 1000, 800],
      ];
      items.forEach((it, i) => {
        const r = 5 + i;
        const zebra = i % 2 ? { bg: '#f3faf6' } : {};
        put(s, `A${r}`, it[0], { ...zebra, bd: thin() });
        put(s, `B${r}`, it[1], { ...zebra, bd: thin(), ha: 'center', color: it[1] === 'Receita' ? '#2f9e6e' : '#c92a2a' });
        put(s, `C${r}`, it[2], { ...CUR, ...zebra, bd: thin() });
        put(s, `D${r}`, it[3], { ...CUR, ...zebra, bd: thin() });
        put(s, `E${r}`, `=D${r}-C${r}`, { ...CUR, ...zebra, bd: thin() });
        put(s, `F${r}`, `=SEERRO(D${r}/C${r};0)`, { fmt: 'percent', dec: 0, ...zebra, bd: thin() });
      });
      const last = 4 + items.length;
      const tot = last + 2;
      put(s, `A${tot}`, 'Total de receitas', { b: true });
      put(s, `D${tot}`, `=SOMASE(B5:B${last};"Receita";D5:D${last})`, { ...CUR, b: true, color: '#2f9e6e' });
      put(s, `A${tot + 1}`, 'Total de despesas', { b: true });
      put(s, `D${tot + 1}`, `=SOMASE(B5:B${last};"Despesa";D5:D${last})`, { ...CUR, b: true, color: '#c92a2a' });
      put(s, `A${tot + 2}`, 'Saldo do mês', { b: true, size: 13 });
      put(s, `D${tot + 2}`, `=D${tot}-D${tot + 1}`, { ...CUR, b: true, size: 13, bd: { t: { w: 'medium', color: '#1d2129' } } });
      s.cols.setSize(0, 230); s.cols.setSize(1, 100); s.cols.setSize(2, 120); s.cols.setSize(3, 120); s.cols.setSize(4, 120); s.cols.setSize(5, 110);
      s.rows.setSize(0, 34); s.rows.setSize(3, 30);
      s.freeze = { r: 4, c: 0 };
      s.cf.push({ id: 'cf1', type: 'bar', range: { r1: 4, c1: 3, r2: last - 1, c2: 3 }, params: { color: '#3fcf8e' } });
      s.cf.push({ id: 'cf2', type: 'lt', range: { r1: 4, c1: 4, r2: last - 1, c2: 4 }, params: { a: 0 }, style: { color: '#c92a2a', bg: '#fff0f0' } });
      s.charts.push({ id: 'ch1', type: 'doughnut', title: 'Despesas por categoria', range: { r1: 6, c1: 0, r2: last - 1, c2: 0 }, seriesIn: 'cols', x: 800, y: 70, w: 460, h: 330, legend: 'right', sheetId: s.id, dataCols: [0, 3] });
      // gráfico usa colunas A e D: guardamos um intervalo auxiliar
      s.charts[0].range = { r1: 6, c1: 0, r2: last - 1, c2: 3 };
      s.charts[0].pick = [0, 3];
      s.charts[0].header = false;
      wb.recalc();
      return wb;
    },
  },

  tasks: {
    title: 'Lista de tarefas',
    build() {
      const wb = new Workbook();
      const s = wb.sheet;
      s.name = 'Tarefas';
      put(s, 'A1', 'Lista de tarefas', TITLE);
      put(s, 'A2', '="Concluídas: "&CONT.SE(E5:E30;"Concluído")&" de "&CONT.VALORES(A5:A30)', SUB);
      row(s, 4, ['Tarefa', 'Responsável', 'Prazo', 'Prioridade', 'Status', 'Dias restantes'], { ...HEAD, bg: '#3b6cf6' });
      const today = new Date();
      const d = (n) => serialFromParts(today.getFullYear(), today.getMonth() + 1, today.getDate() + n);
      const tasks = [
        ['Definir escopo do projeto', 'Ana', d(-3), 'Alta', 'Concluído'], ['Levantar orçamento', 'Bruno', d(2), 'Alta', 'Em andamento'],
        ['Criar protótipo', 'Carla', d(6), 'Média', 'Em andamento'], ['Validar com clientes', 'Ana', d(10), 'Média', 'A fazer'],
        ['Preparar apresentação', 'Diego', d(-1), 'Alta', 'Atrasado'], ['Revisar contrato', 'Bruno', d(14), 'Baixa', 'A fazer'],
      ];
      tasks.forEach((t, i) => {
        const r = 5 + i;
        put(s, `A${r}`, t[0], { bd: thin() });
        put(s, `B${r}`, t[1], { bd: thin() });
        put(s, `C${r}`, t[2], { fmt: 'date', bd: thin(), ha: 'center' });
        put(s, `D${r}`, t[3], { bd: thin(), ha: 'center' });
        put(s, `E${r}`, t[4], { bd: thin(), ha: 'center' });
        put(s, `F${r}`, `=SE(E${r}="Concluído";"—";C${r}-HOJE())`, { bd: thin(), ha: 'center' });
      });
      s.cols.setSize(0, 260); s.cols.setSize(1, 120); s.cols.setSize(2, 110); s.cols.setSize(3, 100); s.cols.setSize(4, 130); s.cols.setSize(5, 120);
      s.rows.setSize(0, 34); s.rows.setSize(3, 30);
      s.freeze = { r: 4, c: 1 };
      const R = (c1, c2 = c1) => ({ r1: 4, c1, r2: 29, c2 });
      s.cf.push({ id: 'a', type: 'text', range: R(4), params: { a: 'Concluído' }, style: { bg: '#d3f9d8', color: '#2b8a3e' } });
      s.cf.push({ id: 'b', type: 'text', range: R(4), params: { a: 'Atrasado' }, style: { bg: '#ffe3e3', color: '#c92a2a' } });
      s.cf.push({ id: 'c', type: 'text', range: R(4), params: { a: 'andamento' }, style: { bg: '#fff3bf', color: '#8a6d00' } });
      s.cf.push({ id: 'd', type: 'text', range: R(3), params: { a: 'Alta' }, style: { color: '#c92a2a', bold: true } });
      s.cf.push({ id: 'e', type: 'lt', range: R(5), params: { a: 0 }, style: { color: '#c92a2a', bold: true } });
      s.validations.push({ range: R(4), type: 'list', items: ['A fazer', 'Em andamento', 'Concluído', 'Atrasado'] });
      s.validations.push({ range: R(3), type: 'list', items: ['Alta', 'Média', 'Baixa'] });
      s.filter = { r1: 3, c1: 0, r2: 4 + tasks.length, c2: 5, criteria: {} };
      wb.recalc();
      return wb;
    },
  },

  sales: {
    title: 'Controle de vendas',
    build() {
      const wb = new Workbook();
      const s = wb.sheet;
      s.name = 'Vendas';
      row(s, 1, ['Data', 'Vendedor', 'Produto', 'Qtd', 'Preço', 'Total'], { ...HEAD, bg: '#f0603f' });
      const sellers = ['Ana', 'Bruno', 'Carla', 'Diego'];
      const products = [['Notebook', 4200], ['Monitor', 1100], ['Teclado', 240], ['Mouse', 120], ['Headset', 380]];
      let seed = 7;
      const rnd = () => { seed = (seed * 9301 + 49297) % 233280; return seed / 233280; };
      const n = 24;
      for (let i = 0; i < n; i++) {
        const r = 2 + i;
        const [p, price] = products[Math.floor(rnd() * products.length)];
        put(s, `A${r}`, serialFromParts(2026, 1 + Math.floor(i / 8), 1 + Math.floor(rnd() * 27)), { fmt: 'date', ha: 'center' });
        put(s, `B${r}`, sellers[Math.floor(rnd() * sellers.length)]);
        put(s, `C${r}`, p);
        put(s, `D${r}`, 1 + Math.floor(rnd() * 5), { ha: 'center' });
        put(s, `E${r}`, price, CUR);
        put(s, `F${r}`, `=D${r}*E${r}`, CUR);
      }
      const last = n + 1;
      put(s, 'H1', 'Vendedor', { ...HEAD, bg: '#1d2129' });
      put(s, 'I1', 'Total vendido', { ...HEAD, bg: '#1d2129' });
      put(s, 'J1', 'Vendas', { ...HEAD, bg: '#1d2129' });
      put(s, 'H2', `=CLASSIFICAR(ÚNICO(B2:B${last}))`, { b: true });
      put(s, 'I2', `=SOMASE(B2:B${last};H2#;F2:F${last})`, CUR);
      put(s, 'J2', `=CONT.SE(B2:B${last};H2#)`, { ha: 'center' });
      put(s, 'H7', 'Total geral', { b: true, bd: { t: { w: 'medium' } } });
      put(s, 'I7', `=SOMA(F2:F${last})`, { ...CUR, b: true, bd: { t: { w: 'medium' } } });
      put(s, 'H9', 'Ticket médio', { b: true });
      put(s, 'I9', `=MÉDIA(F2:F${last})`, { ...CUR, b: true });
      put(s, 'H10', 'Maior venda', { b: true });
      put(s, 'I10', `=MÁXIMO(F2:F${last})`, { ...CUR, b: true });
      [90, 90, 110, 50, 110, 120, 24, 120, 130, 70].forEach((w, i) => s.cols.setSize(i, w));
      s.rows.setSize(0, 30);
      s.freeze = { r: 1, c: 0 };
      s.filter = { r1: 0, c1: 0, r2: last - 1, c2: 5, criteria: {} };
      s.cf.push({ id: 's1', type: 'scale3', range: { r1: 1, c1: 5, r2: last - 1, c2: 5 }, params: {} });
      s.charts.push({ id: 'ch1', type: 'column', title: 'Vendas por vendedor', range: { r1: 0, c1: 7, r2: 4, c2: 8 }, seriesIn: 'cols', x: 1020, y: 30, w: 480, h: 300, legend: 'none', sheetId: s.id });
      wb.recalc();
      return wb;
    },
  },
};
