import test from 'node:test'
import assert from 'node:assert/strict'
import { HOME_RIBBON_GROUP_CONFIG } from '../src/utils/homeRibbonConfig.js'

test('The native Home tab and its standard workbook formatting toolbar are available', () => {
  assert.deepEqual(HOME_RIBBON_GROUP_CONFIG, {
    'ribbon.start': { title: 'Home', order: 1, hidden: false },
  })
})

test('No duplicate Home groups or custom toolbar commands replace native controls', async () => {
  const { readFile } = await import('node:fs/promises')
  const source = await readFile(new URL('../src/components/UniverEditor.jsx', import.meta.url), 'utf8')

  assert.doesNotMatch(source, /const homeGroups = \[/)
  assert.doesNotMatch(source, /const homeGroupMenus =/)
  assert.doesNotMatch(source, /\.\.\.homeGroupMenus/)
  assert.doesNotMatch(source, /gridnova\.ribbon\.home\./)
})

test('Home remains available in both initial preset setup and runtime menu updates', async () => {
  const { readFile } = await import('node:fs/promises')
  const source = await readFile(new URL('../src/components/UniverEditor.jsx', import.meta.url), 'utf8')

  assert.match(source, /menu:\s*\{\s*\.\.\.HOME_RIBBON_GROUP_CONFIG/s)
  assert.match(source, /univerAPI\.updateMenuConfig\(\{\s*\.\.\.HOME_RIBBON_GROUP_CONFIG/s)
})

test('Other workbook tabs remain available alongside Home', async () => {
  const { readFile } = await import('node:fs/promises')
  const source = await readFile(new URL('../src/components/UniverEditor.jsx', import.meta.url), 'utf8')

  assert.match(source, /'ribbon\.insert':\s*\{\s*title:\s*'Insert',[^}]*hidden:\s*false/s)
  assert.match(source, /'ribbon\.formulas':\s*\{\s*title:\s*'Formulas'/)
  assert.match(source, /'ribbon\.data':\s*\{\s*title:\s*'Data'/)
  assert.match(source, /'ribbon\.view':\s*\{\s*title:\s*'View'/)
})

test('The editor keeps all non-Home native ribbon tabs and menu groups', async () => {
  const { readFile } = await import('node:fs/promises')
  const source = await readFile(new URL('../src/components/UniverEditor.jsx', import.meta.url), 'utf8')

  assert.match(source, /const nativeTabGroups = \[/)
  assert.match(source, /path: 'ribbon\.insert'/)
  assert.match(source, /path: 'ribbon\.data'/)
  assert.match(source, /return \[\.\.\.icons, \.\.\.drawingComponents, \.\.\.menus, \.\.\.nativeTabGroupMenus\]/)
})
