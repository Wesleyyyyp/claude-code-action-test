const mysql = require('mysql');

// Look up a user by the username/password they typed in
function login(db, username, password) {
  const query = "SELECT * FROM users WHERE username = '" + username + "' AND password = '" + password + "'";
  return db.query(query, function (err, results) {
    if (err) throw err;
    if (results.length > 0) {
      console.log("Login successful for password: " + password);
      return results[0];
    }
    return null;
  });
}

module.exports = { login };
