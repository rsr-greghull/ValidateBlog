using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.ServiceModel.Syndication;

public class SyndicatedItem {
    public string Title;
    public List<string> Labels;
    public Uri url;
    public string Body;
    public SyndicatedItem() { }
}

public class Syndicator {
    SyndicationFeed Feed;
    Dictionary<Uri, SyndicatedItem> BloggerData;
    public int Errors { get { return _Errors; } }
    int _Errors = 0;
    public Syndicator(string fileName) {
    XmlReader reader = XmlReader.Create(fileName);
        Feed = SyndicationFeed.Load(reader);
        reader.Close();
        BloggerData = new Dictionary<Uri, SyndicatedItem>(2000);
    }
    // Make unknown/other the default value for all these enums
    enum Kind { unknown, comment, page, post, settings, template };
    enum Label { Other, Ratings, Review, Blog };
    enum Category { Unknown, Short_Story, Novelette, Novella, Novel, Anthology, Collection };

    public Dictionary<Uri, SyndicatedItem> Validate(bool nonReviewFlag, bool verbose) {

        int comments = 0;
        int settings = 0;
        int templates = 0;
        int pages = 0;
        int posts = 0;
        int unknowns = 0;
        int labels = 0;
        int ratings = 0;
        int reviews = 0;
        int blogs = 0;
        int others = 0;
        int anthologies = 0;
        int collections = 0;
        int novels = 0;
        int novelettes = 0;
        int novellas = 0;
        int short_stories = 0;
        int unknownCats = 0;
        int postDrafts = 0;
        int pageDrafts = 0;
        int[] ratingStars = new int[6];

        // Process every item in the feed--some are posts, s
        foreach (SyndicationItem item in Feed.Items) {
            string subject = item.Title.Text;

            // See if it's a draft
            // This will call non-page and post items "drafts" but we'll fix that later.
            bool isDraft = true;

            Uri url = null;
            if (subject != null && subject.Length > 0) {
                foreach (SyndicationLink link in item.Links) {
                    string title = link.Title;
                    switch (title) {
                    case null:
                    case "":
                    case "Post Comments":
                        break;
                    default:
                        if (title.Length > 0 && title == subject) {
                            isDraft = false;
                            url = link.Uri;
                            if (url.Host != "www.rocketstackrank.com") {
                                Console.WriteLine("Unexpected Host: '{0}'", url.Host);
                                ++_Errors;
                            }
                            break;
                        } else {
                            string[] fields = title.Split(' ');
                            if (fields.Length == 2 && fields[1] == "Comments") {
                                break;
                            }
                        }
                        Console.WriteLine("Unexpected Link Title {0}", link.Title);
                        ++_Errors;
                        break;
                    }
                }
            }

            // Find out what sort of entry we've got
            Kind kind = Kind.unknown;
            Label label = Label.Other;
            Category lengthCat = Category.Unknown; // e.g. novella, novelette etc.
            bool isRating = false;
            bool isReview = false;
            bool isBlog = false;

            bool isAnthology = false;
            bool isCollection = false;
            bool isNovel = false;
            bool isNovelette = false;
            bool isNovella = false;
            bool isShort_Story = false;

            bool hasRating = false;
            int rating = 0;
            int categories = 0;
            int catLabels = 0;
            var labelList = new List<string>(10); // List of all labels found on post
            foreach (SyndicationCategory category in item.Categories) {
                string scheme = category.Scheme;
                string name = category.Name;
                if (scheme == "http://schemas.google.com/g/2005#kind") {
                    string[] fields = name.Split('#');
                    if (fields[0] == "http://schemas.google.com/blogger/2008/kind") {
                        if (!Enum.TryParse<Kind>(fields[1], out kind)) {
                            Console.WriteLine("Unknown Kind: {0}", fields[1]);
                            ++_Errors;
                        }
                        switch (kind) {
                        case Kind.comment:
                            comments++;
                            break;
                        case Kind.page:
                            pages++;
                            break;
                        case Kind.post:
                            posts++;
                            break;
                        case Kind.settings:
                            settings++;
                            break;
                        case Kind.template:
                            templates++;
                            break;
                        default:
                            unknowns++;
                            break;
                        }
                    } else {
                        Console.WriteLine("Unknown Name: {0}", name);
                        ++_Errors;
                    }
                } else if (scheme == "http://www.blogger.com/atom/ns#") {
                    labelList.Add(name);
                    Enum.TryParse<Label>(name, out label);
                    switch (label) {
                    case Label.Ratings:
                        isRating = true;
                        ++ratings;
                        ++catLabels;
                        break;
                    case Label.Review:
                        isReview = true;
                        ++reviews;
                        ++catLabels;
                        break;
                    case Label.Blog:
                        isBlog = true;
                        ++blogs;
                        ++catLabels;
                        break;
                    case Label.Other:
                        ++others;
                        ++catLabels;
                        break;
                    }
                    ++labels;

                    Enum.TryParse<Category>(name, out lengthCat);
                    if (name == "Short Story") {
                        lengthCat = Category.Short_Story;
                    }
                    if (!isDraft) {
                        switch (lengthCat) {
                        case Category.Anthology:
                            isAnthology = true;
                            anthologies++;
                            categories++;
                            break;
                        case Category.Collection:
                            isCollection = true;
                            categories++;
                            collections++;
                            break;
                        case Category.Novel:
                            isNovel = true;
                            novels++;
                            categories++;
                            break;
                        case Category.Novelette:
                            isNovelette = true;
                            novelettes++;
                            categories++;
                            break;
                        case Category.Novella:
                            isNovella = true;
                            novellas++;
                            categories++;
                            break;
                        case Category.Short_Story:
                            isShort_Story = true;
                            short_stories++;
                            categories++;
                            break;
                        }
                    }

                    string[] fields = name.Split(':');
                    if (fields[0] == "Rating") {
                        if (hasRating) {
                            Console.WriteLine("Two Ratings for one review! {0}\t{1}\t{2}", scheme, category.Label, category.Name);
                            ++_Errors;
                        }
                        hasRating = true;
                        if (fields[1] == " NR") {
                            rating = 0;
                        } else {
                            rating = int.Parse(fields[1]);
                        }
                        ratingStars[rating]++;
                    }

                } else {
                    Console.WriteLine("Unknown Scheme: {0}\t{1}\t{2}", scheme, category.Label, category.Name);
                    ++_Errors;
                }
            }

            if (kind != Kind.post) {
                if (catLabels > 0) {
                    Console.WriteLine("Label(s) on a non-post!");
                    Console.WriteLine("{0}\t{1}", kind, subject);
                    ++_Errors;
                }
            }

            // Not really an error, but we'd like to know
            if (isDraft) {
                if (catLabels > 0) {
                    Console.WriteLine("Label(s) on a draft");
                    Console.WriteLine("{0}\t{1}", kind, subject);
                }
                if (kind == Kind.page) {
                    pageDrafts++;
                }
                if (kind == Kind.post) {
                    postDrafts++;
                }
            }

            // Don't waste time analyzing drafts
            if (isDraft) {
                continue;
            }

            // Special case for announcements
            if (isReview && isBlog && isRating) {
                isReview = isRating = false;
            }
            if (isReview) {
                if (categories == 0) {
                    Console.WriteLine("Review has no category!");
                    Console.WriteLine("{0}\t{1}", kind, subject);
                    ++_Errors;
                }
                if (categories > 1) {
                    Console.WriteLine("Review has {0} categories!", categories);
                    Console.WriteLine("{0}\t{1}", kind, subject);
                    ++_Errors;
                }

                if (!hasRating) {
                    Console.WriteLine("Review has no rating!");
                    Console.WriteLine("{0}\t{1}", kind, subject);
                    ++_Errors;
                }

                if (isBlog) {
                    Console.WriteLine("Review is a blog!");
                    Console.WriteLine("{0}\t{1}", kind, subject);
                    ++_Errors;
                }

                if (isAnthology || isCollection) {
                    if (!isRating) {
                        Console.WriteLine("Anthology/Collection Review is not a Rating!");
                        Console.WriteLine("{0}\t{1}", kind, subject);
                        ++_Errors;
                    }

                } else {
                    if (isRating) {
                        Console.WriteLine("Non Anthology/Collection Review is a Rating!");
                        Console.WriteLine("{0}\t{1}", kind, subject);
                        ++_Errors;
                    }
                }
                SyndicatedItem synItem = new SyndicatedItem();
                synItem.Title = subject;
                synItem.Labels = labelList;
                synItem.Body = ((TextSyndicationContent)item.Content).Text;
                synItem.url = url;
                if (BloggerData.ContainsKey(url)) {
                    Console.WriteLine("Syndicated Item already exists at {0}!", url);
                    Console.WriteLine("{0}\t{1}", kind, subject);
                    ++_Errors;
                } else {
                    BloggerData.Add(url, synItem);
                }
            } else {
                // Just report this one--these are not errors
                if (kind == Kind.post && !isDraft && nonReviewFlag && !isBlog && !isRating) {
                    Console.WriteLine("{0}\t{1}\t{2}", kind, subject, url.AbsoluteUri);
                }
                if (categories != 0) {
                    Console.WriteLine("Non-Review has {0} categories!", categories);
                    Console.WriteLine("{0}\t{1}", kind, subject);
                    ++_Errors;
                }
                if (hasRating) {
                    Console.WriteLine("Non-Review has rating {0}!", rating);
                    Console.WriteLine("{0}\t{1}", kind, subject);
                    ++_Errors;
                }
            }

            TextSyndicationContent content = (TextSyndicationContent)item.Content;

        }
        if (verbose) {
            Console.WriteLine("Templates:\t{0}", templates);
            Console.WriteLine("Settings:\t{0}", settings);
            Console.WriteLine("Pages:\t{0}", pages);
            Console.WriteLine("Posts:\t{0}", posts);
            Console.WriteLine("Comments:\t{0}", comments);
            Console.WriteLine("Labels:\t{0}", labels);
            Console.WriteLine(" Ratings:\t{0}", ratings);
            Console.WriteLine(" Reviews:\t{0}", reviews);
            Console.WriteLine("  Short Stories:\t{0}", short_stories);
            Console.WriteLine("  Novelettes:\t{0}", novelettes);
            Console.WriteLine("  Novellas:\t{0}", novellas);
            Console.WriteLine("  Novels:\t{0}", novels);
            Console.WriteLine("  Anthologies:\t{0}", anthologies);
            Console.WriteLine("  Collections:\t{0}", collections);
            Console.WriteLine("  Unknown:\t{0}", unknownCats);
            Console.WriteLine(" Blogs:\t{0}", blogs);
            Console.WriteLine(" Others:\t{0}", others);
            Console.WriteLine("Post Drafts:\t{0}", postDrafts);
            Console.WriteLine("Page Drafts:\t{0}", pageDrafts);
            for (int i = 0; i< 6; ++i) {
                Console.WriteLine("Rating: {0}\t{1}", i, ratingStars[i]);
            }
            Console.WriteLine();
        }
        return BloggerData;
    }
}
