require('dotenv').config();
const {
  Client,
  GatewayIntentBits,
  Events,
  SlashCommandBuilder,
  PermissionFlagsBits,
  REST,
  Routes,
} = require('discord.js');
const { google } = require('googleapis');
const config = require('./roster.config.json');

const {
  DISCORD_TOKEN,
  GUILD_ID,
  SHEET_ID,
  SHEET_TAB = 'Staff Roster',
  GOOGLE_KEY_FILE = './service-account.json',
  REFRESH_MINUTES = '10',
} = process.env;

for (const [k, v] of Object.entries({ DISCORD_TOKEN, GUILD_ID, SHEET_ID })) {
  if (!v) {
    console.error(`Missing ${k} in .env`);
    process.exit(1);
  }
}

const client = new Client({
  intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMembers],
});

const auth = new google.auth.GoogleAuth({
  keyFile: GOOGLE_KEY_FILE,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

const HEADERS = ['Rank', 'Name', 'Username', 'User ID', 'Joined'];
const COLS = HEADERS.length;


const byName = (a, b) =>
  a.displayName.localeCompare(b.displayName, undefined, { sensitivity: 'base' });

async function buildRoster(guild) {
  await guild.members.fetch();

  const isEligible = (m) => config.showBots || !m.user.bot;
  const assigned = new Set(); 
  const sections = [];

  for (const { id, name } of config.staffRoles) {
    const role = guild.roles.cache.get(id);
    if (!role) {
     
      console.warn(`Role ID ${id} ("${name}") not found in the server. Check the ID in roster.config.json.`);
      sections.push({ title: name, color: 0, members: [] });
      continue;
    }
    const members = guild.members.cache
      .filter((m) => isEligible(m) && !assigned.has(m.id) && m.roles.cache.has(id))
      .sort(byName);
    members.forEach((m) => assigned.add(m.id));
    sections.push({ title: name, color: role.color, members: [...members.values()] });
  }

  if (config.includeMembers) {
    const memberCfg = config.memberRole && config.memberRole.id ? config.memberRole : null;
    const members = guild.members.cache
      .filter(
        (m) =>
          isEligible(m) &&
          !assigned.has(m.id) &&
          (!memberCfg || m.roles.cache.has(memberCfg.id))
      )
      .sort(byName);
    sections.push({
      title: memberCfg ? memberCfg.name : 'Members',
      color: 0,
      members: [...members.values()],
    });
  }

  return sections;
}


function toRows(sections) {
  const rows = [];
  const headerRows = [];

  const stamp = new Date().toLocaleString('en-US', {
    timeZone: 'America/New_York',
    dateStyle: 'medium',
    timeStyle: 'short',
  });
  rows.push([`Staff Roster — last updated ${stamp} ET`, '', '', '', '']);
  rows.push(['', '', '', '', '']);
  rows.push(HEADERS);

  for (const s of sections) {
    headerRows.push({ index: rows.length, color: s.color });
    rows.push([`${s.title} (${s.members.length})`, '', '', '', '']);
    if (s.members.length === 0) {
      rows.push([s.title, '(vacant)', '', '', '']);
      continue;
    }
    for (const m of s.members) {
      rows.push([
        s.title,
        m.displayName,
        m.user.username,
        m.id,
        m.joinedAt ? m.joinedAt.toISOString().slice(0, 10) : '',
      ]);
    }
  }
  return { rows, headerRows };
}

const intToRgb = (n) =>
  n === 0
    ? { red: 0.85, green: 0.85, blue: 0.85 }
    : {
        red: ((n >> 16) & 255) / 255,
        green: ((n >> 8) & 255) / 255,
        blue: (n & 255) / 255,
      };

const isDark = ({ red, green, blue }) => 0.299 * red + 0.587 * green + 0.114 * blue < 0.55;



async function ensureTab() {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  let tab = meta.data.sheets.find((s) => s.properties.title === SHEET_TAB);
  if (!tab) {
    const res = await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_TAB } } }] },
    });
    return res.data.replies[0].addSheet.properties.sheetId;
  }
  return tab.properties.sheetId;
}

async function writeSheet(rows, headerRows) {
  const sheetId = await ensureTab();
  const range = `'${SHEET_TAB}'`;


  await sheets.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range });


  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: `${range}!A1`,
    valueInputOption: 'RAW',
    requestBody: { values: rows },
  });


  const full = { sheetId };
  const row = (r) => ({ sheetId, startRowIndex: r, endRowIndex: r + 1, startColumnIndex: 0, endColumnIndex: COLS });

  const requests = [

    { repeatCell: { range: full, cell: {}, fields: 'userEnteredFormat' } },
  
    {
      repeatCell: {
        range: row(0),
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 14 } } },
        fields: 'userEnteredFormat.textFormat',
      },
    },

    {
      repeatCell: {
        range: row(2),
        cell: {
          userEnteredFormat: {
            textFormat: { bold: true },
            backgroundColor: { red: 0.93, green: 0.93, blue: 0.93 },
          },
        },
        fields: 'userEnteredFormat(textFormat,backgroundColor)',
      },
    },

    {
      updateSheetProperties: {
        properties: { sheetId, gridProperties: { frozenRowCount: 3 } },
        fields: 'gridProperties.frozenRowCount',
      },
    },
  ];


  for (const h of headerRows) {
    const bg = intToRgb(h.color);
    const fg = isDark(bg) ? { red: 1, green: 1, blue: 1 } : { red: 0, green: 0, blue: 0 };
    requests.push({
      repeatCell: {
        range: row(h.index),
        cell: {
          userEnteredFormat: {
            backgroundColor: bg,
            textFormat: { bold: true, foregroundColor: fg },
          },
        },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
  }

  requests.push({
    autoResizeDimensions: {
      dimensions: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: COLS },
    },
  });

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    requestBody: { requests },
  });
}



let running = false;
let queued = false;

async function syncRoster(reason = 'manual') {
  if (running) {
    queued = true;
    return;
  }
  running = true;
  try {
    const guild = await client.guilds.fetch(GUILD_ID);
    const sections = await buildRoster(guild);
    const { rows, headerRows } = toRows(sections);
    await writeSheet(rows, headerRows);
    const total = sections.reduce((n, s) => n + s.members.length, 0);
    console.log(`[${new Date().toISOString()}] Roster synced (${reason}): ${total} people.`);
  } catch (err) {
    console.error(`Roster sync failed (${reason}):`, err.message || err);
  } finally {
    running = false;
    if (queued) {
      queued = false;
      syncRoster('queued');
    }
  }
}


let timer = null;
function scheduleSync(reason) {
  clearTimeout(timer);
  timer = setTimeout(() => syncRoster(reason), 5000);
}



const syncCommand = new SlashCommandBuilder()
  .setName('syncroster')
  .setDescription('Force the staff roster Google Sheet to update now')
  .setDefaultMemberPermissions(PermissionFlagsBits.ManageGuild);

client.once(Events.ClientReady, async () => {
  console.log(`Logged in as ${client.user.tag}`);

  try {
    const rest = new REST({ version: '10' }).setToken(DISCORD_TOKEN);
    await rest.put(Routes.applicationGuildCommands(client.user.id, GUILD_ID), {
      body: [syncCommand.toJSON()],
    });
  } catch (err) {
    console.error('Could not register /syncroster:', err.message);
  }

  await syncRoster('startup');
  setInterval(() => syncRoster('interval'), Math.max(1, Number(REFRESH_MINUTES)) * 60 * 1000);
});

client.on(Events.GuildMemberAdd, (m) => m.guild.id === GUILD_ID && scheduleSync('member joined'));
client.on(Events.GuildMemberRemove, (m) => m.guild.id === GUILD_ID && scheduleSync('member left'));
client.on(Events.GuildMemberUpdate, (oldM, newM) => {
  if (newM.guild.id !== GUILD_ID) return;
  const rolesChanged =
    oldM.roles.cache.size !== newM.roles.cache.size ||
    oldM.roles.cache.some((_, id) => !newM.roles.cache.has(id));
  const nameChanged = oldM.displayName !== newM.displayName;
  if (rolesChanged || nameChanged) scheduleSync('member updated');
});
client.on(Events.GuildRoleUpdate, (_o, r) => r.guild.id === GUILD_ID && scheduleSync('role updated'));
client.on(Events.GuildRoleDelete, (r) => r.guild.id === GUILD_ID && scheduleSync('role deleted'));

client.on(Events.InteractionCreate, async (interaction) => {
  if (!interaction.isChatInputCommand() || interaction.commandName !== 'syncroster') return;
  await interaction.reply({ content: 'Updating the roster sheet…', ephemeral: true });
  await syncRoster(`command by ${interaction.user.tag}`);
  await interaction.editReply('Roster sheet updated.');
});

process.on('unhandledRejection', (err) => console.error('Unhandled rejection:', err));

client.login(DISCORD_TOKEN);
