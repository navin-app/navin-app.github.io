// ============================================================
//  NAVIN SOCIAL PLATFORM — Google Apps Script Backend v2
//  Features: Social Feed, Coins, Groups, Friends, Hydration
// ============================================================

const SPREADSHEET_ID = '1aHZ-GpFx-uM94Sbycf6r5b0JELqb07NxdSBpsb1AWrg';
const SHEET_USERS    = 'Users';
const SHEET_POSTS    = 'Posts';
const SHEET_FRIENDS  = 'Friends';
const SHEET_GROUPS   = 'Groups';
const SHEET_EVENTS   = 'Events';

// ── Ensure sheets exist ────────────────────────────────────
function ensureSheetsExist() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let users = ss.getSheetByName(SHEET_USERS);
  if (!users) {
    users = ss.insertSheet(SHEET_USERS);
    users.appendRow(['userId','email','passwordHash','name','joinDate','lastLogin','level','coins','totalHydration_L','avatarBase64']);
    users.getRange(1,1,1,10).setFontWeight('bold');
  } else {
    const headers = users.getRange(1, 1, 1, users.getLastColumn()).getValues()[0];
    if (!headers.includes('level')) {
      users.getRange(1, users.getLastColumn() + 1).setValue('level');
      users.getRange(1, users.getLastColumn() + 1).setValue('coins');
      users.getRange(1, users.getLastColumn() + 1).setValue('totalHydration_L');
    }
    if (!headers.includes('avatarBase64')) {
      users.getRange(1, users.getLastColumn() + 1).setValue('avatarBase64');
    }
  }

  let posts = ss.getSheetByName(SHEET_POSTS);
  if (!posts) {
    posts = ss.insertSheet(SHEET_POSTS);
    posts.appendRow(['postId','userId','email','name','photoBase64','caption','timestamp','coins_earned','likes_count','likedBy']);
    posts.getRange(1,1,1,10).setFontWeight('bold');
  }

  let friends = ss.getSheetByName(SHEET_FRIENDS);
  if (!friends) {
    friends = ss.insertSheet(SHEET_FRIENDS);
    friends.appendRow(['userId','friendId','status','createdAt']);
    friends.getRange(1,1,1,4).setFontWeight('bold');
  }

  let groups = ss.getSheetByName(SHEET_GROUPS);
  if (!groups) {
    groups = ss.insertSheet(SHEET_GROUPS);
    groups.appendRow(['groupId','creatorId','name','members','createdAt']);
    groups.getRange(1,1,1,5).setFontWeight('bold');
  }

  let events = ss.getSheetByName(SHEET_EVENTS);
  if (!events) {
    events = ss.insertSheet(SHEET_EVENTS);
    events.appendRow(['eventId','groupId','title','date','hashtag','createdAt']);
    events.getRange(1,1,1,6).setFontWeight('bold');
  }
  } catch (err) {
    Logger.log('Sheet initialization error: ' + err.message);
  }
}

// ── Process Request ────────────────────────────────────────
function processRequest(payload) {
  ensureSheetsExist();

  const action = payload.action;
  let result;

  if      (action === 'ping')           result = { success: true, message: 'NAVIN v2 online' };
  else if (action === 'signup')         result = handleSignup(payload);
  else if (action === 'login')          result = handleLogin(payload);
  else if (action === 'getMyProfile')   result = handleGetMyProfile(payload);
  else if (action === 'updateProfile')  result = handleUpdateProfile(payload);
  else if (action === 'getFeed')        result = handleGetFeed(payload);
  else if (action === 'uploadPost')     result = handleUploadPost(payload);
  else if (action === 'deletePost')     result = handleDeletePost(payload);
  else if (action === 'likePost')       result = handleLikePost(payload);
  else if (action === 'getMyFriends')   result = handleGetMyFriends(payload);
  else if (action === 'addFriend')      result = handleAddFriend(payload);
  else if (action === 'addFriendByUserId') result = handleAddFriendByUserId(payload);
  else if (action === 'updateHydration')result = handleUpdateHydration(payload);
  else if (action === 'getGroups')      result = handleGetGroups(payload);
  else if (action === 'createGroup')    result = handleCreateGroup(payload);
  else                                  result = { success: false, message: 'Unknown action: ' + action };

  return result;
}

// ── doGet ──────────────────────────────────────────────────
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const payload = e.parameter && e.parameter.payload
      ? JSON.parse(decodeURIComponent(e.parameter.payload))
      : { action: 'ping' };

    const result = processRequest(payload);
    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, message: 'Server error: ' + err.message }));
  }

  return output;
}

// ── doPost ─────────────────────────────────────────────────
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    let payload = {};

    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.payload) {
      payload = JSON.parse(decodeURIComponent(e.parameter.payload));
    }

    const result = processRequest(payload);
    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, message: 'Server error: ' + err.message }));
  }

  return output;
}

// ── Signup ─────────────────────────────────────────────────
function handleSignup(body) {
  const { email, password, name } = body;
  if (!email || !password || !name)
    return { success: false, message: 'Email, password, dan nama wajib diisi.' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) {
    ensureSheetsExist();
    return handleSignup(body);
  }

  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().toLowerCase() === email.toLowerCase())
      return { success: false, message: 'Email sudah terdaftar.' };
  }

  const userId       = 'USR_' + Utilities.getUuid().substring(0, 8).toUpperCase();
  const passwordHash = hashPassword(password);
  const joinDate     = new Date().toISOString();

  sheet.appendRow([userId, email.toLowerCase(), passwordHash, name, joinDate, joinDate, 'Movers', 0, 0, '']);

  const token = generateToken(userId, email);
  return { success: true, message: 'Signup berhasil!', token, userId, name, email: email.toLowerCase() };
}

// ── Login ──────────────────────────────────────────────────
function handleLogin(body) {
  const { email, password } = body;
  if (!email || !password)
    return { success: false, message: 'Email dan password wajib diisi.' };

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return { success: false, message: 'Belum ada user.' };

  const data  = sheet.getDataRange().getValues();
  const passwordHash = hashPassword(password);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] && row[1].toString().toLowerCase() === email.toLowerCase() && row[2] === passwordHash) {
      sheet.getRange(i + 1, 6).setValue(new Date().toISOString());
      const token = generateToken(row[0], row[1]);
      return {
        success: true,
        message: 'Login berhasil!',
        token,
        userId: row[0],
        name: row[3],
        email: row[1],
        level: row[6] || 'Movers',
        coins: row[7] || 0,
        avatarBase64: row[9] || ''
      };
    }
  }

  return { success: false, message: 'Email atau password salah.' };
}

// ── Get My Profile ────────────────────────────────────────
function handleGetMyProfile(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const usersData = usersSheet.getDataRange().getValues();

  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) {
      return {
        success: true,
        profile: {
          userId: usersData[i][0],
          email: usersData[i][1],
          name: usersData[i][3],
          joinDate: usersData[i][4],
          level: usersData[i][6] || 'Movers',
          coins: usersData[i][7] || 0,
          totalHydration_L: usersData[i][8] || 0,
          avatarBase64: usersData[i][9] || ''
        }
      };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

// ── Update Profile ────────────────────────────────────────
function handleUpdateProfile(body) {
  const { token, userId, name, avatarBase64 } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const usersData = usersSheet.getDataRange().getValues();

  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) {
      if (name) usersSheet.getRange(i + 1, 4).setValue(name);
      if (avatarBase64) usersSheet.getRange(i + 1, 10).setValue(avatarBase64.substring(0, 40000));
      return { success: true, message: 'Profil diupdate!' };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

// ── Upload Post ───────────────────────────────────────────
function handleUploadPost(body) {
  const { token, userId, caption, photoBase64 } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const postsSheet = ss.getSheetByName(SHEET_POSTS);
  const usersSheet = ss.getSheetByName(SHEET_USERS);

  const usersData = usersSheet.getDataRange().getValues();
  let userName = '', userEmail = '';
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) {
      userName = usersData[i][3];
      userEmail = usersData[i][1];
      break;
    }
  }

  const postId = 'POST_' + new Date().getTime();
  const photoRef = photoBase64 && photoBase64.length > 0 ? photoBase64.substring(0, 40000) : '';

  postsSheet.appendRow([
    postId,
    userId,
    userEmail,
    userName,
    photoRef,
    caption,
    new Date().toISOString(),
    20,
    0,
    ''
  ]);

  const usersSheet2 = ss.getSheetByName(SHEET_USERS);
  const usersData2 = usersSheet2.getDataRange().getValues();
  for (let i = 1; i < usersData2.length; i++) {
    if (usersData2[i][0] === userId) {
      const currentCoins = usersData2[i][7] || 0;
      usersSheet2.getRange(i + 1, 8).setValue(currentCoins + 20);
      break;
    }
  }

  return { success: true, message: 'Post berhasil dibuat! +20 coins', postId };
}

// ── Delete Post ────────────────────────────────────────────
function handleDeletePost(body) {
  const { token, userId, postId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const postsSheet = ss.getSheetByName(SHEET_POSTS);
  const postsData = postsSheet.getDataRange().getValues();

  for (let i = 1; i < postsData.length; i++) {
    if (postsData[i][0] === postId) {
      if (postsData[i][1] !== userId) {
        return { success: false, message: 'Hanya pemilik post yang bisa menghapus.' };
      }
      postsSheet.deleteRow(i + 1);
      return { success: true, message: 'Post dihapus.' };
    }
  }

  return { success: false, message: 'Post tidak ditemukan.' };
}

// ── Like Post ─────────────────────────────────────────────
function handleLikePost(body) {
  const { token, userId, postId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const postsSheet = ss.getSheetByName(SHEET_POSTS);
  const postsData = postsSheet.getDataRange().getValues();

  for (let i = 1; i < postsData.length; i++) {
    if (postsData[i][0] === postId) {
      const likes = (postsData[i][8] || 0) + 1;
      const likedBy = (postsData[i][9] || '').length > 0 ? postsData[i][9] + ',' + userId : userId;

      postsSheet.getRange(i + 1, 9).setValue(likes);
      postsSheet.getRange(i + 1, 10).setValue(likedBy);

      const creatorId = postsData[i][1];
      const usersSheet = ss.getSheetByName(SHEET_USERS);
      const usersData = usersSheet.getDataRange().getValues();
      for (let j = 1; j < usersData.length; j++) {
        if (usersData[j][0] === creatorId) {
          const currentCoins = usersData[j][7] || 0;
          usersSheet.getRange(j + 1, 8).setValue(currentCoins + 1);
          break;
        }
      }

      return { success: true, message: 'Post dilike! +1 coin untuk author' };
    }
  }

  return { success: false, message: 'Post tidak ditemukan.' };
}

// ── Get Feed (Friends' Posts) ─────────────────────────────
function handleGetFeed(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const friendsSheet = ss.getSheetByName(SHEET_FRIENDS);
  const postsSheet = ss.getSheetByName(SHEET_POSTS);

  const friendsData = friendsSheet.getDataRange().getValues();
  const friendIds = [userId];
  for (let i = 1; i < friendsData.length; i++) {
    if (friendsData[i][0] === userId && friendsData[i][2] === 'accepted') {
      friendIds.push(friendsData[i][1]);
    }
  }

  const postsData = postsSheet.getDataRange().getValues();
  const feed = [];
  for (let i = postsData.length - 1; i >= 1; i--) {
    if (friendIds.includes(postsData[i][1])) {
      feed.push({
        postId: postsData[i][0],
        userId: postsData[i][1],
        name: postsData[i][3],
        email: postsData[i][2],
        photoBase64: postsData[i][4] || '',
        caption: postsData[i][5] || '',
        timestamp: postsData[i][6],
        coins_earned: postsData[i][7] || 0,
        likes_count: postsData[i][8] || 0,
        likedBy: postsData[i][9] || ''
      });
    }
  }

  return { success: true, feed: feed.slice(0, 50) };
}

// ── Get My Friends ────────────────────────────────────────
function handleGetMyFriends(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const friendsSheet = ss.getSheetByName(SHEET_FRIENDS);
  const usersSheet = ss.getSheetByName(SHEET_USERS);

  const friendsData = friendsSheet.getDataRange().getValues();
  const usersData = usersSheet.getDataRange().getValues();

  const userMap = {};
  usersData.slice(1).forEach(row => {
    userMap[row[0]] = { name: row[3], email: row[1], level: row[6] || 'Movers' };
  });

  const friends = [];
  for (let i = 1; i < friendsData.length; i++) {
    if (friendsData[i][0] === userId && friendsData[i][2] === 'accepted') {
      const friendId = friendsData[i][1];
      friends.push({
        friendId: friendId,
        name: userMap[friendId]?.name || 'Unknown',
        email: userMap[friendId]?.email || friendId,
        level: userMap[friendId]?.level || 'Movers'
      });
    }
  }

  return { success: true, friends };
}

// ── Add Friend by Email ───────────────────────────────────
function handleAddFriend(body) {
  const { token, userId, friendEmail } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const friendsSheet = ss.getSheetByName(SHEET_FRIENDS);

  const usersData = usersSheet.getDataRange().getValues();
  let friendId = null;
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][1] === friendEmail.toLowerCase()) {
      friendId = usersData[i][0];
      break;
    }
  }

  if (!friendId) return { success: false, message: 'Email tidak ditemukan.' };

  friendsSheet.appendRow([userId, friendId, 'accepted', new Date().toISOString()]);
  friendsSheet.appendRow([friendId, userId, 'accepted', new Date().toISOString()]);

  return { success: true, message: 'Friend ditambahkan!' };
}

// ── Add Friend by User ID (privacy-friendly) ──────────────
function handleAddFriendByUserId(body) {
  const { token, userId, friendUserId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const friendsSheet = ss.getSheetByName(SHEET_FRIENDS);

  const usersData = usersSheet.getDataRange().getValues();
  let friendExists = false;
  let friendName = '';
  let resolvedFriendId = friendUserId;
  const query = (friendUserId || '').toString().trim().toLowerCase();

  for (let i = 1; i < usersData.length; i++) {
    const rowUserId = (usersData[i][0] || '').toString();
    const rowEmail = (usersData[i][1] || '').toString().toLowerCase();
    const rowName = (usersData[i][3] || '').toString().toLowerCase();

    if (rowUserId === friendUserId || rowEmail === query || rowName === query) {
      friendExists = true;
      friendName = usersData[i][3];
      resolvedFriendId = rowUserId;
      break;
    }
  }

  if (!friendExists) return { success: false, message: 'User tidak ditemukan. Coba pakai User ID, email, atau nama lengkap yang persis.' };
  if (userId === resolvedFriendId) return { success: false, message: 'Tidak bisa add diri sendiri.' };

  const friendsData = friendsSheet.getDataRange().getValues();
  for (let i = 1; i < friendsData.length; i++) {
    if ((friendsData[i][0] === userId && friendsData[i][1] === resolvedFriendId) ||
        (friendsData[i][0] === resolvedFriendId && friendsData[i][1] === userId)) {
      return { success: false, message: 'Sudah berteman.' };
    }
  }

  friendsSheet.appendRow([userId, resolvedFriendId, 'accepted', new Date().toISOString()]);
  friendsSheet.appendRow([resolvedFriendId, userId, 'accepted', new Date().toISOString()]);

  return { success: true, message: 'Teman ditambahkan: ' + friendName };
}

// ── Update Hydration ──────────────────────────────────────
function handleUpdateHydration(body) {
  const { token, userId, hydration_ml } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const usersData = usersSheet.getDataRange().getValues();

  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) {
      const currentHydration = usersData[i][8] || 0;
      const newHydration = currentHydration + (hydration_ml / 1000);
      const dailyGoalReached = currentHydration < 1 && newHydration >= 1;

      usersSheet.getRange(i + 1, 9).setValue(newHydration);

      let coinsBonus = 0;
      if (dailyGoalReached) {
        const currentCoins = usersData[i][7] || 0;
        usersSheet.getRange(i + 1, 8).setValue(currentCoins + 5);
        coinsBonus = 5;
      }

      return {
        success: true,
        message: 'Hydration updated!',
        totalHydration: Math.round(newHydration * 100) / 100,
        dailyGoalReached,
        coinsBonus
      };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

// ── Get Groups (AOC only) ──────────────────────────────────
function handleGetGroups(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const groupsSheet = ss.getSheetByName(SHEET_GROUPS);
  const usersSheet = ss.getSheetByName(SHEET_USERS);

  const usersData = usersSheet.getDataRange().getValues();
  let userLevel = 'Movers';
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) {
      userLevel = usersData[i][6] || 'Movers';
      break;
    }
  }

  if (!groupsSheet) return { success: true, groups: [], canCreate: userLevel.toUpperCase() === 'AOC' };

  const groupsData = groupsSheet.getDataRange().getValues();
  const groups = [];
  for (let i = 1; i < groupsData.length; i++) {
    groups.push({
      groupId: groupsData[i][0],
      creatorId: groupsData[i][1],
      name: groupsData[i][2],
      members: (groupsData[i][3] || '').split(','),
      createdAt: groupsData[i][4]
    });
  }

  return { success: true, groups, canCreate: userLevel.toUpperCase() === 'AOC' };
}

// ── Create Group (AOC only) ────────────────────────────────
function handleCreateGroup(body) {
  const { token, userId, groupName } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const groupsSheet = ss.getSheetByName(SHEET_GROUPS);

  const usersData = usersSheet.getDataRange().getValues();
  let userLevel = 'Movers';
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) {
      userLevel = usersData[i][6] || 'Movers';
      break;
    }
  }

  if (userLevel.toUpperCase() !== 'AOC') {
    return { success: false, message: 'Hanya AOC yang bisa membuat group.' };
  }

  const groupId = 'GRP_' + new Date().getTime();
  groupsSheet.appendRow([groupId, userId, groupName, userId, new Date().toISOString()]);

  return { success: true, message: 'Grup dibuat!', groupId };
}

// ── Helpers ───────────────────────────────────────────────
function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + 'NAVIN_SALT_2026');
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function generateToken(userId, email) {
  const payload = userId + '|' + email + '|' + new Date().getTime();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload + 'NAVIN_TOKEN_SECRET');
  const hash = digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('').substring(0, 32);
  return Utilities.base64Encode(userId + ':' + hash);
}

function verifyToken(token, userId) {
  if (!token || !userId) return false;
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    return decoded.startsWith(userId + ':');
  } catch (e) {
    return false;
  }
}
