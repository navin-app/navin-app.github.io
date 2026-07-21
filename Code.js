// ============================================================
//  NAVIN SOCIAL PLATFORM — Google Apps Script Backend v2
//  Features: Social Feed, Coins, Groups, Friends, Hydration
// ============================================================

const SPREADSHEET_ID = '1aHZ-GpFx-uM94Sbycf6r5b0JELqb07NxdSBpsb1AWrg';
const SHEET_USERS     = 'Users';
const SHEET_POSTS     = 'Posts';
const SHEET_FRIENDS   = 'Friends';
const SHEET_GROUPS    = 'Groups';
const SHEET_EVENTS    = 'Events';
const SHEET_CHATS     = 'Chats';
const SHEET_REFERRALS = 'ReferralCodes';
const SHEET_REF_USAGE = 'ReferralUsage';
const SHEET_INBOX     = 'Inbox';
const SHEET_ADS       = 'Ads';

// Poin
const PTS_HYDRATION_SUBMIT = 25;  // Mover/AoC submit 500ml dengan kode valid
const PTS_REFERRAL_OWNER   = 2;   // AoC pemilik kode, tiap kode dipakai

// ── Ensure sheets exist ────────────────────────────────────
function ensureSheetsExist() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let users = ss.getSheetByName(SHEET_USERS);
  if (!users) {
    users = ss.insertSheet(SHEET_USERS);
    users.appendRow(['userId','email','passwordHash','name','joinDate','lastLogin','level','coins','totalHydration_L','avatarBase64','hydrationToken']);
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
    if (!headers.includes('hydrationToken')) {
      users.getRange(1, users.getLastColumn() + 1).setValue('hydrationToken');
    }
  }

  let posts = ss.getSheetByName(SHEET_POSTS);
  if (!posts) {
    posts = ss.insertSheet(SHEET_POSTS);
    posts.appendRow(['postId','userId','email','name','photoBase64','caption','timestamp','coins_earned','likes_count','likedBy','activity','location','distance_km','duration_s']);
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

  let chats = ss.getSheetByName(SHEET_CHATS);
  if (!chats) {
    chats = ss.insertSheet(SHEET_CHATS);
    chats.appendRow(['chatKey','chatType','senderId','senderName','message','timestamp']);
    chats.getRange(1,1,1,6).setFontWeight('bold');
  }

  let referrals = ss.getSheetByName(SHEET_REFERRALS);
  if (!referrals) {
    referrals = ss.insertSheet(SHEET_REFERRALS);
    referrals.appendRow(['code','ownerId','ownerName','status','durationDays','requestedAt','approvedAt','expiresAt','note']);
    referrals.getRange(1,1,1,9).setFontWeight('bold');
  }

  let refUsage = ss.getSheetByName(SHEET_REF_USAGE);
  if (!refUsage) {
    refUsage = ss.insertSheet(SHEET_REF_USAGE);
    refUsage.appendRow(['code','usedBy','usedByName','ownerId','usedAt']);
    refUsage.getRange(1,1,1,5).setFontWeight('bold');
  }

  let inbox = ss.getSheetByName(SHEET_INBOX);
  if (!inbox) {
    inbox = ss.insertSheet(SHEET_INBOX);
    inbox.appendRow(['inboxId','userId','type','title','message','read','createdAt']);
    inbox.getRange(1,1,1,7).setFontWeight('bold');
  }

  let ads = ss.getSheetByName(SHEET_ADS);
  if (!ads) {
    ads = ss.insertSheet(SHEET_ADS);
    ads.appendRow(['adId','imageBase64','title','linkUrl','uploadedBy','createdAt']);
    ads.getRange(1,1,1,6).setFontWeight('bold');
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
  else if (action === 'uploadChunk')    result = handleUploadChunk(payload);
  else if (action === 'uploadPost')     result = handleUploadPost(payload);
  else if (action === 'deletePost')     result = handleDeletePost(payload);
  else if (action === 'likePost')       result = handleLikePost(payload);
  else if (action === 'getMyFriends')   result = handleGetMyFriends(payload);
  else if (action === 'addFriend')      result = handleAddFriend(payload);
  else if (action === 'addFriendByUserId') result = handleAddFriendByUserId(payload);
  else if (action === 'updateHydration')result = handleUpdateHydration(payload);
  else if (action === 'getGroups')      result = handleGetGroups(payload);
  else if (action === 'createGroup')    result = handleCreateGroup(payload);
  else if (action === 'inviteToGroup')  result = handleInviteToGroup(payload);
  else if (action === 'getGroupMembers')result = handleGetGroupMembers(payload);
  else if (action === 'sendChat')       result = handleSendChat(payload);
  else if (action === 'getChat')        result = handleGetChat(payload);
  else if (action === 'requestReferralCode') result = handleRequestReferralCode(payload);
  else if (action === 'getMyReferralCodes')  result = handleGetMyReferralCodes(payload);
  else if (action === 'getPendingReferrals') result = handleGetPendingReferrals(payload);
  else if (action === 'approveReferralCode') result = handleApproveReferralCode(payload);
  else if (action === 'rejectReferralCode')  result = handleRejectReferralCode(payload);
  else if (action === 'submitHydrationCode') result = handleSubmitHydrationCode(payload);
  else if (action === 'getInbox')       result = handleGetInbox(payload);
  else if (action === 'markInboxRead')  result = handleMarkInboxRead(payload);
  else if (action === 'getAds')         result = handleGetAds(payload);
  else if (action === 'uploadAd')       result = handleUploadAd(payload);
  else if (action === 'deleteAd')       result = handleDeleteAd(payload);
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
          avatarBase64: usersData[i][9] || '',
          hydrationToken: usersData[i][10] || ''
        }
      };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

// ── Update Profile ────────────────────────────────────────
function handleUpdateProfile(body) {
  const { token, userId, name, avatarBase64, avatarUploadId, avatarChunks } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEET_USERS);
  const usersData = usersSheet.getDataRange().getValues();

  let avatar = avatarBase64 || '';
  if (avatarUploadId && avatarChunks) {
    const assembled = assembleChunks(avatarUploadId, avatarChunks);
    if (assembled === null) return { success: false, message: 'Upload avatar tidak lengkap, coba lagi.' };
    avatar = assembled;
  }

  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) {
      if (name) usersSheet.getRange(i + 1, 4).setValue(name);
      if (avatar) usersSheet.getRange(i + 1, 10).setValue(avatar.substring(0, 45000));
      return { success: true, message: 'Profil diupdate!' };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

// ── Chunked Upload (untuk foto > limit URL GET) ────────────
function handleUploadChunk(body) {
  const { token, userId, uploadId, chunkIndex, totalChunks, data } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };
  if (!uploadId || chunkIndex === undefined || !data)
    return { success: false, message: 'Chunk data tidak lengkap.' };

  const cache = CacheService.getScriptCache();
  cache.put('upl_' + uploadId + '_' + chunkIndex, data, 3600);
  return { success: true, message: 'Chunk ' + (chunkIndex + 1) + '/' + totalChunks + ' diterima' };
}

function assembleChunks(uploadId, totalChunks) {
  const cache = CacheService.getScriptCache();
  const keys = [];
  for (let i = 0; i < totalChunks; i++) keys.push('upl_' + uploadId + '_' + i);
  const got = cache.getAll(keys);
  let out = '';
  for (let i = 0; i < totalChunks; i++) {
    const part = got['upl_' + uploadId + '_' + i];
    if (part == null) return null;
    out += part;
  }
  cache.removeAll(keys);
  return out;
}

// ── Upload Post ───────────────────────────────────────────
function handleUploadPost(body) {
  const { token, userId, caption, photoBase64, photoUploadId, photoChunks, activity, location, distance_km, duration_s } = body;
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

  // Foto: dari chunks (upload besar) atau langsung (kecil)
  let photo = photoBase64 || '';
  if (photoUploadId && photoChunks) {
    const assembled = assembleChunks(photoUploadId, photoChunks);
    if (assembled === null) return { success: false, message: 'Upload foto tidak lengkap, coba lagi.' };
    photo = assembled;
  }

  const postId = 'POST_' + new Date().getTime();
  const photoRef = photo && photo.length > 0 ? photo.substring(0, 45000) : '';

  // ── Quest: poin aktivitas dari durasi ──
  // >= 5 menit baru dapat poin. Kelipatan tiap 5 menit = +1 poin. Cap 3 jam (36 poin).
  const durSec = parseInt(duration_s) || 0;
  let movePoints = 0;
  if (durSec >= 300) {
    movePoints = Math.min(Math.floor(durSec / 300), 36);
  }
  const POST_BASE = 20;
  const totalEarned = POST_BASE + movePoints;

  postsSheet.appendRow([
    postId,
    userId,
    userEmail,
    userName,
    photoRef,
    caption,
    new Date().toISOString(),
    totalEarned,
    0,
    '',
    activity || '',
    location || '',
    distance_km || '',
    duration_s || ''
  ]);

  const usersSheet2 = ss.getSheetByName(SHEET_USERS);
  const usersData2 = usersSheet2.getDataRange().getValues();
  for (let i = 1; i < usersData2.length; i++) {
    if (usersData2[i][0] === userId) {
      const currentCoins = usersData2[i][7] || 0;
      usersSheet2.getRange(i + 1, 8).setValue(currentCoins + totalEarned);
      break;
    }
  }

  let msg = 'Post berhasil dibuat! +' + POST_BASE + ' coins';
  if (movePoints > 0) {
    msg += ' + ' + movePoints + ' Quest Points (' + Math.floor(durSec / 60) + ' menit ' + (activity || 'aktivitas') + ')';
  } else if (durSec > 0 && durSec < 300) {
    msg += ' (aktivitas < 5 menit belum dapat Quest Point)';
  }

  return { success: true, message: msg, postId, movePoints, coinsEarned: totalEarned };
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
      // Cek: user sudah pernah like? Tidak boleh double (tidak ada unlike)
      const likedByStr = (postsData[i][9] || '').toString();
      const likedByArr = likedByStr.length > 0 ? likedByStr.split(',') : [];
      if (likedByArr.indexOf(userId) !== -1) {
        return { success: false, message: 'Kamu sudah like post ini.', alreadyLiked: true };
      }

      const likes = (postsData[i][8] || 0) + 1;
      const likedBy = likedByArr.concat(userId).join(',');
      postsSheet.getRange(i + 1, 9).setValue(likes);
      postsSheet.getRange(i + 1, 10).setValue(likedBy);

      // +1 poin untuk yang MELAKUKAN like (bukan author)
      const usersSheet = ss.getSheetByName(SHEET_USERS);
      const usersData = usersSheet.getDataRange().getValues();
      for (let j = 1; j < usersData.length; j++) {
        if (usersData[j][0] === userId) {
          const currentCoins = usersData[j][7] || 0;
          usersSheet.getRange(j + 1, 8).setValue(currentCoins + 1);
          break;
        }
      }

      return { success: true, message: 'Post dilike! +1 poin', likes };
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
  for (let i = postsData.length - 1; i >= 1 && feed.length < 20; i--) {
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
        likedBy: postsData[i][9] || '',
        activity: postsData[i][10] || '',
        location: postsData[i][11] || '',
        distance_km: postsData[i][12] || '',
        duration_s: postsData[i][13] || ''
      });
    }
  }

  return { success: true, feed };
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

      // Token validasi hidrasi: dibuat SEKALI per profil, permanen
      let hydrationToken = usersData[i][10] || '';
      let newToken = false;
      if (!hydrationToken) {
        hydrationToken = 'NAVIN-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
        usersSheet.getRange(i + 1, 11).setValue(hydrationToken);
        newToken = true;
      }

      return {
        success: true,
        message: 'Hydration updated!',
        totalHydration: Math.round(newHydration * 100) / 100,
        dailyGoalReached,
        coinsBonus,
        hydrationToken,
        newToken
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
    const creatorId = groupsData[i][1];
    const members = (groupsData[i][3] || '').split(',').filter(m => m.length > 0);
    // Hanya tampilkan group jika user adalah creator ATAU member
    if (creatorId !== userId && !members.includes(userId)) continue;
    groups.push({
      groupId: groupsData[i][0],
      creatorId: creatorId,
      name: groupsData[i][2],
      members: members,
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

// ── Invite Friend to Group (AOC creator only) ──────────────
function handleInviteToGroup(body) {
  const { token, userId, groupId, friendUserId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };
  if (!friendUserId) return { success: false, message: 'Target friend tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const groupsSheet = ss.getSheetByName(SHEET_GROUPS);
  if (!groupsSheet) return { success: false, message: 'Grup tidak ditemukan.' };

  const groupsData = groupsSheet.getDataRange().getValues();
  let groupRow = -1, groupMembers = [];
  for (let i = 1; i < groupsData.length; i++) {
    if (groupsData[i][0] === groupId) {
      if (groupsData[i][1] !== userId) {
        return { success: false, message: 'Hanya pembuat grup yang bisa invite teman.' };
      }
      groupRow = i + 1;
      groupMembers = (groupsData[i][3] || '').split(',').filter(m => m.length > 0);
      break;
    }
  }

  if (groupRow === -1) return { success: false, message: 'Grup tidak ditemukan.' };
  if (groupMembers.includes(friendUserId)) {
    return { success: false, message: 'User sudah member grup ini.' };
  }

  groupMembers.push(friendUserId);
  groupsSheet.getRange(groupRow, 4).setValue(groupMembers.join(','));

  return { success: true, message: 'Teman ditambahkan ke grup!' };
}

// ── Get Group Members ──────────────────────────────────────
function handleGetGroupMembers(body) {
  const { token, userId, groupId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const groupsSheet = ss.getSheetByName(SHEET_GROUPS);
  const usersSheet = ss.getSheetByName(SHEET_USERS);

  if (!groupsSheet) return { success: true, groupInfo: {}, members: [] };

  const groupsData = groupsSheet.getDataRange().getValues();
  const usersData = usersSheet.getDataRange().getValues();
  const userMap = {};
  for (let i = 1; i < usersData.length; i++) {
    userMap[usersData[i][0]] = { name: usersData[i][3], level: usersData[i][6] || 'Movers' };
  }

  for (let i = 1; i < groupsData.length; i++) {
    if (groupsData[i][0] === groupId) {
      const creatorId = groupsData[i][1];
      const memberIds = (groupsData[i][3] || '').split(',').filter(m => m.length > 0);
      const members = memberIds.map(mId => ({ userId: mId, name: userMap[mId]?.name || 'Unknown', level: userMap[mId]?.level || 'Movers' }));

      return {
        success: true,
        groupInfo: {
          groupId: groupsData[i][0],
          name: groupsData[i][2],
          creatorId,
          creatorName: userMap[creatorId]?.name || 'Unknown',
          createdAt: groupsData[i][4],
          isCreator: creatorId === userId
        },
        members
      };
    }
  }

  return { success: false, message: 'Grup tidak ditemukan.' };
}

// ── Chat: kunci konsisten untuk friend chat (2 arah sama) ──
function friendChatKey(a, b) {
  return [a, b].sort().join('|');
}

// ── Send Chat Message ──────────────────────────────────────
function handleSendChat(body) {
  const { token, userId, chatType, targetId, message } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };
  if (!message || !message.trim()) return { success: false, message: 'Pesan kosong.' };
  if (!targetId) return { success: false, message: 'Target chat tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const chatsSheet = ss.getSheetByName(SHEET_CHATS);
  const usersSheet = ss.getSheetByName(SHEET_USERS);

  // Ambil nama pengirim
  const usersData = usersSheet.getDataRange().getValues();
  let senderName = 'Unknown';
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) { senderName = usersData[i][3]; break; }
  }

  const type = chatType === 'group' ? 'group' : 'friend';
  const chatKey = type === 'group' ? targetId : friendChatKey(userId, targetId);
  const timestamp = new Date().toISOString();

  chatsSheet.appendRow([chatKey, type, userId, senderName, message.trim().substring(0, 500), timestamp]);

  return { success: true, message: 'Terkirim', sent: { senderId: userId, senderName, message: message.trim().substring(0, 500), timestamp } };
}

// ── Get Chat Messages ──────────────────────────────────────
function handleGetChat(body) {
  const { token, userId, chatType, targetId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };
  if (!targetId) return { success: false, message: 'Target chat tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const chatsSheet = ss.getSheetByName(SHEET_CHATS);
  if (!chatsSheet) return { success: true, messages: [] };

  const type = chatType === 'group' ? 'group' : 'friend';
  const chatKey = type === 'group' ? targetId : friendChatKey(userId, targetId);

  const chatsData = chatsSheet.getDataRange().getValues();
  const messages = [];
  // Scan dari bawah (terbaru), ambil max 50, lalu balik urutan
  for (let i = chatsData.length - 1; i >= 1 && messages.length < 50; i--) {
    if (chatsData[i][0] === chatKey && chatsData[i][1] === type) {
      messages.push({
        senderId: chatsData[i][2],
        senderName: chatsData[i][3],
        message: chatsData[i][4],
        timestamp: chatsData[i][5]
      });
    }
  }
  messages.reverse();

  return { success: true, messages };
}

// ============================================================
//  REFERRAL CODE SYSTEM  (Admin → AoC → Movers)
// ============================================================

function genReferralCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // tanpa 0/O/1/I biar tidak ambigu
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, Utilities.getUuid());
  let s = '';
  for (let i = 0; i < 6; i++) s += chars[(bytes[i] & 0xFF) % chars.length];
  return 'AOC-' + s;
}

function addInbox(userId, type, title, message) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const inbox = ss.getSheetByName(SHEET_INBOX);
  if (!inbox) return;
  inbox.appendRow(['INB_' + Utilities.getUuid().substring(0, 12), userId, type, title, message, false, new Date().toISOString()]);
}

function getUserLevel(usersData, userId) {
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId) return (usersData[i][6] || 'Movers');
  }
  return 'Movers';
}

// ── AoC: Request Referral Code ─────────────────────────────
function handleRequestReferralCode(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersData = ss.getSheetByName(SHEET_USERS).getDataRange().getValues();
  let userName = '';
  for (let i = 1; i < usersData.length; i++) if (usersData[i][0] === userId) userName = usersData[i][3];

  if (getUserLevel(usersData, userId).toUpperCase() !== 'AOC') {
    return { success: false, message: 'Hanya AOC yang bisa request referral code.' };
  }

  const refSheet = ss.getSheetByName(SHEET_REFERRALS);
  const refData = refSheet.getDataRange().getValues();
  for (let i = 1; i < refData.length; i++) {
    if (refData[i][1] === userId && refData[i][3] === 'pending') {
      return { success: false, message: 'Masih ada request pending. Tunggu approval admin dulu.' };
    }
  }

  const code = genReferralCode();
  refSheet.appendRow([code, userId, userName, 'pending', '', new Date().toISOString(), '', '', '']);
  return { success: true, message: 'Request terkirim! Menunggu approval admin.', code };
}

// ── AoC: Get My Referral Codes ─────────────────────────────
function handleGetMyReferralCodes(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const refSheet = ss.getSheetByName(SHEET_REFERRALS);
  if (!refSheet) return { success: true, codes: [] };
  const refData = refSheet.getDataRange().getValues();
  const usageSheet = ss.getSheetByName(SHEET_REF_USAGE);
  const usageData = usageSheet ? usageSheet.getDataRange().getValues() : [];

  const codes = [];
  const now = new Date();
  for (let i = 1; i < refData.length; i++) {
    if (refData[i][1] === userId) {
      let status = refData[i][3];
      const expiresAt = refData[i][7];
      if (status === 'active' && expiresAt && new Date(expiresAt) < now) status = 'expired';
      const code = (refData[i][0] || '').toString().toUpperCase();
      let usedCount = 0;
      for (let j = 1; j < usageData.length; j++) {
        if ((usageData[j][0] || '').toString().toUpperCase() === code) usedCount++;
      }
      codes.push({
        code: refData[i][0], status, durationDays: refData[i][4],
        approvedAt: refData[i][6], expiresAt, usedCount, note: refData[i][8] || ''
      });
    }
  }
  codes.reverse();
  return { success: true, codes };
}

// ── Admin: List Pending Requests ───────────────────────────
function handleGetPendingReferrals(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersData = ss.getSheetByName(SHEET_USERS).getDataRange().getValues();
  if (getUserLevel(usersData, userId).toUpperCase() !== 'ADMIN') {
    return { success: false, message: 'Hanya Admin.' };
  }

  const refData = ss.getSheetByName(SHEET_REFERRALS).getDataRange().getValues();
  const pending = [];
  for (let i = 1; i < refData.length; i++) {
    if (refData[i][3] === 'pending') {
      pending.push({ code: refData[i][0], ownerId: refData[i][1], ownerName: refData[i][2], requestedAt: refData[i][5] });
    }
  }
  return { success: true, pending, isAdmin: true };
}

// ── Admin: Approve Referral Code ───────────────────────────
function handleApproveReferralCode(body) {
  const { token, userId, code, durationHours } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersData = ss.getSheetByName(SHEET_USERS).getDataRange().getValues();
  if (getUserLevel(usersData, userId).toUpperCase() !== 'ADMIN') {
    return { success: false, message: 'Hanya Admin.' };
  }

  // Hanya 2, 4, atau 24 jam
  let hours = parseInt(durationHours) || 24;
  if ([2, 4, 24].indexOf(hours) === -1) hours = 24;

  const refSheet = ss.getSheetByName(SHEET_REFERRALS);
  const refData = refSheet.getDataRange().getValues();
  const target = (code || '').toString().toUpperCase();
  for (let i = 1; i < refData.length; i++) {
    if ((refData[i][0] || '').toString().toUpperCase() === target) {
      const now = new Date();
      const exp = new Date(now.getTime() + hours * 3600 * 1000);
      refSheet.getRange(i + 1, 4).setValue('active');
      refSheet.getRange(i + 1, 5).setValue(hours + ' jam');
      refSheet.getRange(i + 1, 7).setValue(now.toISOString());
      refSheet.getRange(i + 1, 8).setValue(exp.toISOString());
      addInbox(refData[i][1], 'referral_approved', '✅ Referral Code Disetujui',
        'Kode ' + refData[i][0] + ' aktif ' + hours + ' jam (s/d ' + exp.toLocaleString('id-ID') + '). Bagikan ke Movers!');
      return { success: true, message: 'Approved! Kode aktif ' + hours + ' jam.' };
    }
  }
  return { success: false, message: 'Kode tidak ditemukan.' };
}

// ── Admin: Reject Referral Code ────────────────────────────
function handleRejectReferralCode(body) {
  const { token, userId, code, note } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersData = ss.getSheetByName(SHEET_USERS).getDataRange().getValues();
  if (getUserLevel(usersData, userId).toUpperCase() !== 'ADMIN') {
    return { success: false, message: 'Hanya Admin.' };
  }

  const refSheet = ss.getSheetByName(SHEET_REFERRALS);
  const refData = refSheet.getDataRange().getValues();
  const target = (code || '').toString().toUpperCase();
  for (let i = 1; i < refData.length; i++) {
    if ((refData[i][0] || '').toString().toUpperCase() === target) {
      refSheet.getRange(i + 1, 4).setValue('rejected');
      refSheet.getRange(i + 1, 9).setValue(note || '');
      addInbox(refData[i][1], 'referral_rejected', '❌ Referral Code Ditolak',
        'Kode ' + refData[i][0] + ' ditolak admin.' + (note ? ' Alasan: ' + note : ''));
      return { success: true, message: 'Ditolak.' };
    }
  }
  return { success: false, message: 'Kode tidak ditemukan.' };
}

// ── Mover/AoC: Submit 500ml dengan Referral Code ───────────
function handleSubmitHydrationCode(body) {
  const { token, userId, referralCode } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };
  if (!referralCode || !referralCode.trim()) return { success: false, message: 'Masukkan referral code dulu.' };

  const code = referralCode.trim().toUpperCase();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const refSheet = ss.getSheetByName(SHEET_REFERRALS);
  const usageSheet = ss.getSheetByName(SHEET_REF_USAGE);
  const usersSheet = ss.getSheetByName(SHEET_USERS);

  // 1. Cari & validasi kode
  const refData = refSheet.getDataRange().getValues();
  let codeRow = -1, ownerId = '', status = '', expiresAt = null;
  for (let i = 1; i < refData.length; i++) {
    if ((refData[i][0] || '').toString().toUpperCase() === code) {
      codeRow = i; ownerId = refData[i][1]; status = refData[i][3]; expiresAt = refData[i][7];
      break;
    }
  }
  if (codeRow === -1) return { success: false, message: 'Referral code tidak ditemukan.' };
  if (status === 'pending')  return { success: false, message: 'Kode belum di-approve admin.' };
  if (status === 'rejected') return { success: false, message: 'Kode ini ditolak admin.' };
  if (status !== 'active')   return { success: false, message: 'Kode tidak aktif (status: ' + status + ').' };
  if (expiresAt && new Date(expiresAt) < new Date()) {
    refSheet.getRange(codeRow + 1, 4).setValue('expired');
    addInbox(ownerId, 'referral_expired', '⌛ Referral Code Expired', 'Kode ' + code + ' sudah kedaluwarsa.');
    return { success: false, message: 'Referral code sudah expired.' };
  }

  // 2. Cek pemakaian: 1 kode hanya 1x per profil
  const usageData = usageSheet.getDataRange().getValues();
  for (let i = 1; i < usageData.length; i++) {
    if ((usageData[i][0] || '').toString().toUpperCase() === code && usageData[i][1] === userId) {
      return { success: false, message: 'Kamu sudah pakai kode ini. 1 kode hanya bisa dipakai 1x per profil.' };
    }
  }

  // 3. Cari index user & owner
  const usersData = usersSheet.getDataRange().getValues();
  let userIdx = -1, ownerIdx = -1, userName = '';
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] === userId)  { userIdx = i; userName = usersData[i][3]; }
    if (usersData[i][0] === ownerId) ownerIdx = i;
  }
  if (userIdx === -1) return { success: false, message: 'User tidak ditemukan.' };

  // 4. Tambah hidrasi 0.5L
  const curHydration = usersData[userIdx][8] || 0;
  const newHydration = curHydration + 0.5;
  usersSheet.getRange(userIdx + 1, 9).setValue(newHydration);

  // 5. Poin
  if (ownerIdx === userIdx) {
    // AoC pakai kode sendiri: +25 (submit) +2 (owner)
    const c = usersData[userIdx][7] || 0;
    usersSheet.getRange(userIdx + 1, 8).setValue(c + PTS_HYDRATION_SUBMIT + PTS_REFERRAL_OWNER);
  } else {
    const cu = usersData[userIdx][7] || 0;
    usersSheet.getRange(userIdx + 1, 8).setValue(cu + PTS_HYDRATION_SUBMIT);
    if (ownerIdx !== -1) {
      const co = usersData[ownerIdx][7] || 0;
      usersSheet.getRange(ownerIdx + 1, 8).setValue(co + PTS_REFERRAL_OWNER);
    }
  }

  // 6. Catat pemakaian
  usageSheet.appendRow([code, userId, userName, ownerId, new Date().toISOString()]);

  // 7. Inbox receipt ke user + notif ke owner
  addInbox(userId, 'hydration_receipt', '💧 Bukti Hidrasi',
    'Submit 500ml tervalidasi dengan kode ' + code + '. +' + PTS_HYDRATION_SUBMIT + ' poin! Total hidrasi: ' + (Math.round(newHydration * 100) / 100) + 'L');
  if (ownerId && ownerId !== userId) {
    addInbox(ownerId, 'referral_used', '🎉 Referral Code Dipakai',
      userName + ' memakai kode ' + code + '. +' + PTS_REFERRAL_OWNER + ' poin untuk kamu!');
  }

  // 8. Hydration token (fitur lama, 1 per profil)
  let hydrationToken = usersData[userIdx][10] || '';
  if (!hydrationToken) {
    hydrationToken = 'NAVIN-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
    usersSheet.getRange(userIdx + 1, 11).setValue(hydrationToken);
  }

  return {
    success: true,
    message: 'Hidrasi tervalidasi! +' + PTS_HYDRATION_SUBMIT + ' poin',
    coinsEarned: PTS_HYDRATION_SUBMIT,
    totalHydration: Math.round(newHydration * 100) / 100,
    hydrationToken
  };
}

// ── Get Inbox ──────────────────────────────────────────────
function handleGetInbox(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const inboxSheet = ss.getSheetByName(SHEET_INBOX);
  if (!inboxSheet) return { success: true, inbox: [], unread: 0 };
  const data = inboxSheet.getDataRange().getValues();
  const items = [];
  let unread = 0;
  for (let i = data.length - 1; i >= 1 && items.length < 50; i--) {
    if (data[i][1] === userId) {
      const read = data[i][5] === true || data[i][5] === 'TRUE' || data[i][5] === 'true';
      if (!read) unread++;
      items.push({ inboxId: data[i][0], type: data[i][2], title: data[i][3], message: data[i][4], read, createdAt: data[i][6] });
    }
  }
  return { success: true, inbox: items, unread };
}

// ── Mark Inbox Read ────────────────────────────────────────
function handleMarkInboxRead(body) {
  const { token, userId, inboxId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const inboxSheet = ss.getSheetByName(SHEET_INBOX);
  if (!inboxSheet) return { success: true };
  const data = inboxSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === userId && (!inboxId || data[i][0] === inboxId)) {
      if (data[i][5] !== true) inboxSheet.getRange(i + 1, 6).setValue(true);
    }
  }
  return { success: true };
}

// ============================================================
//  ADS / IKLAN BANNER (Home) — admin managed
// ============================================================

// ── Get Ads (semua user login) ─────────────────────────────
function handleGetAds(body) {
  const { token, userId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const adsSheet = ss.getSheetByName(SHEET_ADS);
  if (!adsSheet) return { success: true, ads: [] };
  const data = adsSheet.getDataRange().getValues();
  const ads = [];
  for (let i = 1; i < data.length; i++) {
    ads.push({
      adId: data[i][0],
      imageBase64: data[i][1] || '',
      title: data[i][2] || '',
      linkUrl: data[i][3] || '',
      createdAt: data[i][5]
    });
  }
  return { success: true, ads };
}

// ── Upload Ad (Admin only, chunked image) ──────────────────
function handleUploadAd(body) {
  const { token, userId, title, linkUrl, imageBase64, imageUploadId, imageChunks } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersData = ss.getSheetByName(SHEET_USERS).getDataRange().getValues();
  if (getUserLevel(usersData, userId).toUpperCase() !== 'ADMIN') {
    return { success: false, message: 'Hanya Admin yang bisa upload iklan.' };
  }

  let image = imageBase64 || '';
  if (imageUploadId && imageChunks) {
    const assembled = assembleChunks(imageUploadId, imageChunks);
    if (assembled === null) return { success: false, message: 'Upload gambar tidak lengkap, coba lagi.' };
    image = assembled;
  }
  if (!image) return { success: false, message: 'Gambar iklan wajib diisi.' };

  const adsSheet = ss.getSheetByName(SHEET_ADS);
  const adId = 'AD_' + new Date().getTime();
  adsSheet.appendRow([adId, image.substring(0, 45000), title || '', linkUrl || '', userId, new Date().toISOString()]);
  return { success: true, message: 'Iklan ditambahkan!', adId };
}

// ── Delete Ad (Admin only) ─────────────────────────────────
function handleDeleteAd(body) {
  const { token, userId, adId } = body;
  if (!verifyToken(token, userId)) return { success: false, message: 'Token tidak valid.' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const usersData = ss.getSheetByName(SHEET_USERS).getDataRange().getValues();
  if (getUserLevel(usersData, userId).toUpperCase() !== 'ADMIN') {
    return { success: false, message: 'Hanya Admin.' };
  }

  const adsSheet = ss.getSheetByName(SHEET_ADS);
  const data = adsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === adId) {
      adsSheet.deleteRow(i + 1);
      return { success: true, message: 'Iklan dihapus.' };
    }
  }
  return { success: false, message: 'Iklan tidak ditemukan.' };
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
