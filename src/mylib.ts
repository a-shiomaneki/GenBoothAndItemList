function concatUrl(pre, post) {
    if (pre.substr(pre.length - 1) != '/') {
        pre = pre + '/';
    };
    if (post.substr(0, 1) == '/') {
        post = post.substr(1, post.length - 1);
    };
    var url = pre + post;
    return url;
};