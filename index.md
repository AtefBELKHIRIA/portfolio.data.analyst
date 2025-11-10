---
layout: default
---

<header>
  <h1>{{ site.title }}</h1>
  <img src="{{ site.profile_image }}" alt="Photo de profil" class="profile">
  <p>{{ site.description }}</p>
</header>

<section>
  {% capture markdown_content %}
  {% include_relative README.md %}
  {% endcapture %}
  {{ markdown_content | markdownify }}
</section>
